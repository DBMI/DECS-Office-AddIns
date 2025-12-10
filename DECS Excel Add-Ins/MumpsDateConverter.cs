using DECS_Excel_Add_Ins.Census;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Security.Cryptography;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Are we dealing with elapsed days or seconds?
     */
    internal enum MumpsDataType
    {
        Days,
        Seconds,
        Unknown
    }

    /**
     * @brief Converts <a href="https://en.wikipedia.org/wiki/MUMPS">MUMPS dates</a> used in Epic to the Excel standard.
     */
    internal class MumpsDateConverter
    {
        // MUMPS reports dates/times in either days or seconds since 12/31/1840.
        // https://en.wikipedia.org/wiki/MUMPS
        // For health records, we'll figure we're looking at dates from 1900 to 2050.
        // So the DAYS constants bracket the these dates in DAYS since 12/31/1840
        // while the SEC constants count the SECONDS since 12/31/1840.
        private const long DAYS_START = 21900;
        private const long DAYS_END = 78000;
        private const long SEC_START = 1890000000;
        private const long SEC_END = 6700000000;
        private MumpsDataType dataType = MumpsDataType.Unknown;

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

                    // If we haven't yet decided whether data are in seconds or minutes,
                    // try to figure it out now.
                    if (dataType == MumpsDataType.Unknown)
                    {
                        dataType = DecideDataType(contents);
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    break;
                }

                Range targetCell = newCol.Cells[rowOffset];
                string formula = "=" + origCell.Address;

                if (dataType == MumpsDataType.Seconds)
                {
                    formula = "=IFERROR(1 + ((" + origCell.Address + "- 1861833600)/86400), \"\")";
                }
                else if (dataType == MumpsDataType.Days)
                {
                    formula = "=IFERROR(1 + (" + origCell.Address + "- 21549), \"\")";
                }

                targetCell.Formula = formula;
            }

            newCol.EntireColumn.NumberFormat = "yyyy-MM-dd HH:mm:ss";
            origCol.Hidden = true;
        }

        /// <summary>
        /// Decides if we're dealing with MUMPS in elapsed seconds or days.
        /// </summary>
        /// <param name="contents">string contents of first cell</param>
        /// <returns>MumpsDataType</returns>
        private MumpsDataType DecideDataType(string contents)
        {
            if (float.TryParse(contents, out float temp))
            {
                if (temp >= DAYS_START && temp <= DAYS_END)
                {
                    return MumpsDataType.Days;
                }

                if (temp >= SEC_START && temp <= SEC_END)
                {
                    return MumpsDataType.Seconds;
                }
            }

            return MumpsDataType.Unknown;
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
