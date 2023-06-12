using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;

namespace DECS_Excel_Add_Ins
{
    internal class ListImporter
    {
        private Application application;
        private readonly string[] IGNORED_WORDS = { "MRN" };
        private const int MAX_LINES_PER_IMPORT = 1000;
        private const string PREAMBLE = "USE [REL_CLARITY];\r\n\r\n";
        private const string SEGMENT_START = "INSERT INTO #MRN_LIST (MRN)\r\nVALUES\r\n";

        public ListImporter()
        {
            this.application = Globals.ThisAddIn.Application;
        }

        private Range GetMrnColumn(Worksheet worksheet, int lastRow)
        {
            Excel.Range column = GetSelectedCol();
            int numDataPointsInSelectedColumn = Utilities.CountCellsWithData(column, lastRow);

            // If selected column is empty, can we just use the first column?
            if (numDataPointsInSelectedColumn <= 1)
            {
                column = worksheet.Columns[1].EntireColumn;
            }

            return column;
        }

        private Range GetSelectedCol()
        {
            Excel.Range rng = (Excel.Range)this.application.Selection;
            Excel.Range selectedColumn = null;

            foreach (Range col in rng.Columns)
            {
                selectedColumn = col;
                break;
            }

            return selectedColumn;
        }

        public void Scan(Worksheet worksheet)
        {
            // Initialize the output .SQL file.
            Workbook workbook = worksheet.Parent;
            string filename = workbook.FullName;

            (StreamWriter writer, string output_filename) = Utilities.OpenOutput(
                input_filename: filename,
                filetype: ".sql"
            );
            writer.Write(PREAMBLE + SEGMENT_START);

            int lastRow = Utilities.FindLastRow(worksheet);
            Range mrnColumn = GetMrnColumn(worksheet, lastRow);

            int lines_written = 0;
            int lines_written_this_chunk = 0;
            Range thisCell;

            for (int rowNumber = 1; rowNumber <= lastRow; rowNumber++)
            {
                thisCell = mrnColumn.Cells[rowNumber];
                string cell_contents;

                try
                {
                    cell_contents = thisCell.Value2.ToString();
                }
                catch
                {
                    // There's nothing in this cell.
                    continue;
                }

                if (string.IsNullOrEmpty(cell_contents))
                    continue;

                // If the line is just "MRN", ignore it.
                if (IGNORED_WORDS.Contains(cell_contents))
                    continue;

                string line_ending;

                writer.Write("('" + cell_contents + "')");
                lines_written++;
                lines_written_this_chunk++;

                if (lines_written < lastRow)
                {
                    if (lines_written_this_chunk < MAX_LINES_PER_IMPORT)
                    {
                        line_ending = ",\r\n";
                    }
                    else
                    {
                        line_ending = ";\r\n\r\n" + SEGMENT_START;
                        lines_written_this_chunk = 0;
                    }
                }
                else
                {
                    line_ending = ";\r\n";
                }

                writer.Write(line_ending);
            }

            writer.Close();
            Utilities.ShowResults(output_filename);
        }
    }
}
