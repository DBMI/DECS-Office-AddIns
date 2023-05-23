using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DECS_Excel_Add_Ins
{
    internal class Utilities
    {
        internal static int CountCellsWithData(Range rng, int lastRow)
        {
            int numCellsWithData = 0;
            Range thisCell;

            for (int rowNumber = 1; rowNumber <= lastRow; rowNumber++)
            {
                thisCell = rng.Cells[rowNumber];
                string cell_contents;

                try
                {
                    cell_contents = thisCell.Value2.ToString();
                }
                catch
                {
                    // There's nothing in this cell.
                    numCellsWithData = rowNumber;
                    break;
                }
            }

            return numCellsWithData;
        }
        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastCol(Worksheet sheet)
        {
            // Unhide All Cells and clear formats
            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();

            // Detect Last used Columns, including cells that contains formulas that result in blank values
            return sheet.UsedRange.Columns.Count;
        }

        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastRow(Worksheet sheet)
        {
            // Unhide All Cells and clear formats
            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();

            // Detect Last used Row, including cells that contains formulas that result in blank values
            return sheet.UsedRange.Rows.Count;
        }

        // Turn the statement of work filename into a .sql filename.
        internal static string FormOutputFilename(string filename, string filetype = ".sql", bool short_version = false)
        {
            string dir = Path.GetDirectoryName(filename);
            string just_the_filename = Path.GetFileNameWithoutExtension(filename);
            string sql_filename = Path.Combine(dir, just_the_filename + filetype);

            if (short_version)
            {
                sql_filename = just_the_filename + filetype;
            }

            return sql_filename;
        }

        internal static List<string> GetColumnNames(Worksheet sheet)
        {
            List<string> names = new List<string>();
            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                names.Add(range.Value.ToString());

                // Move over one column.
                range = range.Offset[0, 1];
            }

            return names;
        }

        // https://stackoverflow.com/q/21219797/18749636
        internal static string GetTimestamp()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        internal static Range InsertnewColumn(Range range, string newColumnName)
        {
            int columnNumber = range.Column;
            Worksheet sheet = range.Worksheet;
            sheet.Columns[columnNumber].Insert();

            Range newRange = range.Offset[0, -1];
            newRange.Value2 = newColumnName;
            return newRange;
        }

        // Open the output StreamWriter object,
        // understanding that we might have to substitute a shorter version of the output filename
        // if the default filename is too long.
        internal static (StreamWriter writer, string opened_filename) OpenOutput(string input_filename, string filetype = ".sql")
        {
            string output_filename = Utilities.FormOutputFilename(filename: input_filename, filetype: filetype, short_version: false);
            StreamWriter writer_obj;

            try
            {
                writer_obj = new StreamWriter(output_filename);
            }
            // https://stackoverflow.com/a/19329123/18749636
            catch (Exception ex) when (
                ex is System.IO.PathTooLongException
                || ex is System.NotSupportedException)
            {
                output_filename = Utilities.FormOutputFilename(filename: input_filename, filetype: filetype, short_version: true);
                writer_obj = new StreamWriter(output_filename);
            }

            return (writer: writer_obj, opened_filename: output_filename);
        }

        // Reassure the user that we've created the desired output file,
        // and display the file once they've seen the message.
        internal static void ShowResults(string output_filename)
        {
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            string message = "Created file '" + output_filename + "'.";
            DialogResult result = MessageBox.Show(message, "Success", buttons);

            if (result == DialogResult.OK)
            {
                Process.Start(output_filename);
            }
        }

        internal static Range TopOfNamedColumn(Worksheet sheet, string columnName)
        {
            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                if (range.Value == columnName)
                {
                    return range;
                }

                // Move over one column.
                range = range.Offset[0, 1];
            }

            return null;
        }
        internal static void WarnColumnNotFound(string columnName)
        {
            string message = "Column '" + columnName + "' not found.";
            string title = "Not Found";
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result = MessageBox.Show(message, title, buttons, MessageBoxIcon.Warning);
        }
    }
}