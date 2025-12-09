using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

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
            Range newCol = Utilities.InsertNewColumn(range: origCol, 
                                                     newColumnName: origName + " converted");
            int rowOffset = 1;

            // Iterate along the rows.
            while (true)
            {
                rowOffset++;

                Range origCell = origCol.Cells[rowOffset];

                // Break once we've found a blank entry.
                try
                {
                    string contents = origCell.Value.ToString();

                    if (string.IsNullOrEmpty(contents))
                    {
                        break;
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    break;
                }

                Range targetCell = newCol.Cells[rowOffset];

                string formula = "=IFERROR(1 + ((" + origCell.Address + "- 1861833600)/86400), \"\")";
                targetCell.Formula = formula;
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
