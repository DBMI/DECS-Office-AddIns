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
    /**
     * @brief Converts <a href="https://en.wikipedia.org/wiki/MUMPS">MUMPS dates</a> used in Epic to the Excel standard.
     */
    internal class MumpsDateConverter
    {
        internal MumpsDateConverter() { }

        /// <summary>
        /// Converts dates in each user-selected column.
        /// </summary>
        /// <param name="worksheet">Active worksheet</param>        
        internal void ConvertColumn(Worksheet worksheet)
        {
            List<int> selectedCols = GetSelectedColumns(worksheet);

            foreach (int col in selectedCols)
            {
                ConvertThisColumn(col, worksheet);
            }
        }

        /// <summary>
        /// Converts dates in this column.
        /// </summary>
        /// <param name="colNum">int number of column to convert</param>
        /// <param name="worksheet">Active worksheet</param>        
        private void ConvertThisColumn(int colNum, Worksheet sheet)
        {
            Range origCol = sheet.Columns[colNum];
            string origName = origCol.Cells[1].Value2;
            Range newCol = Utilities.InsertNewColumn(origCol, origName + " converted");

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

        /// <summary>
        /// Gets the numbers of the selected columns.
        /// </summary>
        /// <param name="worksheet">Active worksheet</param>
        /// <returns>List<int></returns>
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
