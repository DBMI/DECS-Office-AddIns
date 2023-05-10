using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    internal class Utilities
    {
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
        internal static Range InsertNewColumn(Range range, string newColumnName)
        {
            int columnNumber = range.Column;
            Worksheet sheet = range.Worksheet;
            sheet.Columns[columnNumber].Insert();

            Range newRange = range.Offset[0, -1];
            newRange.Value2 = newColumnName;
            return newRange;
        }
        internal static Range TopOfNamedColumn(Worksheet sheet, string columnName)
        {
            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index < lastUsedCol; col_index++)
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
