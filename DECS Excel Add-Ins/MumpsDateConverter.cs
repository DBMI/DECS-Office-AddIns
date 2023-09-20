using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Excel;
using System.Globalization;

namespace DECS_Excel_Add_Ins
{
    internal class MumpsDateConverter
    {
        internal MumpsDateConverter() { }

        internal void ConvertColumn(Worksheet worksheet)
        {
            List<int> selectedCols = GetSelectedColumns(worksheet);

            foreach (int col in selectedCols)
            {
                ConvertThisColumn(col, worksheet);
            }
        }
        private void ConvertThisColumn(int colNum, Worksheet sheet)
        {
            Range origCol = sheet.Columns[colNum];
            Range newCol = AddColumn(origCol);

            Range allRows = Utilities.AllAvailableRows(sheet);

            // Iterate along the rows.
            foreach (Range row in allRows.Rows)
            {
                // The first row is a header.
                if (row.Row > 1)
                {
                    Range origCell = origCol.Cells[row.Row];
                    Range targetCell = newCol.Cells[row.Row];

                    string formula = "=IFERROR(" + origCell.Address + "-21548, \"\")";
                    targetCell.Formula = formula;
                }
            }

            newCol.EntireColumn.NumberFormat = CultureInfo.InvariantCulture.DateTimeFormat.ShortDatePattern;
            origCol.Hidden = true;
        }
        internal Range AddColumn(Range col)
        {
            string origName = col.Cells[1].Value2;
            col.EntireColumn.Insert();

            Worksheet worksheet = col.Worksheet;
            Range newCol = (Range)worksheet.Columns[col.Column - 1];
            newCol.Cells[1].Value2 = origName + " converted";
            return newCol;
        }

        internal Range GetSelectedColumn(Worksheet worksheet)
        {
            Range rng = (Range)worksheet.Application.ActiveCell;

            // Get the selected columns.
            int column = rng.Column;

            Range origCol = (Range)worksheet.Columns[column];
            return origCol;
        }

        internal List<int> GetSelectedColumns(Worksheet worksheet)
        {
            List<int> columns = new List<int>();
            Range rng = (Range)worksheet.Application.Selection;

            foreach(Range col in rng.Columns)
            {
                columns.Add(col.Column);
            }

            return columns;
        }
    }
}
