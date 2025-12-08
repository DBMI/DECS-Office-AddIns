using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Class that formats Excel output to look "nice".
     */
    internal class Formatter
    {
        private const double BOLD_BUMP = 1.05;
        private const double POINTS_PER_CHAR = 1.15;
        private const double MAX_COLUMN_WIDTH = 30.0;

        private int lastColumn;
        private int lastRow;


        internal Formatter() { }

        /// <summary>
        /// Formats all columns to be centered horizontally & vertically.
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        private void CenterlineTheMain(Worksheet worksheet)
        {
            Workbook workbook = worksheet.Application.ActiveWorkbook;

            // Trying to modify while user is editing a cell will result in an error.
            try
            {
                Style style = workbook.Styles.Add("CenteredHeadings");
                style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = XlVAlign.xlVAlignCenter;

                // Only apply to the header row.
                worksheet.Rows[1].Columns.Style = "CenteredHeadings";
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        /// <summary>
        /// Copies formatting from each column in this sheet to the matching column in the next.
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        internal void CopyFormat(Worksheet worksheet)
        {
            lastColumn = worksheet.UsedRange.Columns.Count;
            Range sourceCell = worksheet.Cells[1, 1];

            Worksheet targetSheet = worksheet.Next;

            while (targetSheet != null)
            {
                Range targetCell = targetSheet.Cells[1, 1];

                // 1) Copy each COLUMN's attributes.
                for (int columnOffset = 0; columnOffset < lastColumn; columnOffset++)
                {
                    CopyColumnFormatting(sourceCell.Offset[1, columnOffset],
                                         targetCell.Offset[1, columnOffset]);
                }

                // Copy header ROW attributes.
                // 2) Header text style
                targetCell.EntireRow.Font.Bold = sourceCell.Font.Bold;

                // 3) Font size
                targetCell.EntireRow.Font.Size = sourceCell.Font.Size;

                // 4) Centering
                targetCell.EntireRow.HorizontalAlignment = sourceCell.HorizontalAlignment;
                targetCell.EntireRow.VerticalAlignment = sourceCell.VerticalAlignment;

                // 5) Borders
                // Walk along the columns of the first row until there's no border specified.
                XlLineStyle linestyle = (XlLineStyle) sourceCell.Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
                int colOffset = 0;

                while (linestyle != XlLineStyle.xlLineStyleNone)
                {
                    targetCell.Offset[0, colOffset].Borders[XlBordersIndex.xlEdgeBottom].Color = sourceCell.Offset[0, colOffset].Borders[XlBordersIndex.xlEdgeBottom].Color;
                    targetCell.Offset[0, colOffset].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = sourceCell.Offset[0, colOffset].Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
                    targetCell.Offset[0, colOffset].Borders[XlBordersIndex.xlEdgeBottom].Weight = sourceCell.Offset[0, colOffset].Borders[XlBordersIndex.xlEdgeBottom].Weight;

                    colOffset++;

                    // In case we run off the end of the page.
                    try
                    {
                        linestyle = (XlLineStyle) sourceCell.Offset[0, colOffset].Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        break;
                    }
                    catch (NullReferenceException)
                    {
                        break;
                    }
                }

                targetSheet = targetSheet.Next;
            }
        }

        /// <summary>
        /// Copies formatting from one column to another.
        /// </summary>
        /// <param name="sourceColumn">Where we copy FROM</param>
        /// <param name="targetColumn">Where we copy TO</param>

        internal void CopyColumnFormatting(Range sourceColumn, Range targetColumn)
        {
            // Attributes to copy:

            // 1) Column width
            targetColumn.EntireColumn.ColumnWidth = sourceColumn.Columns.ColumnWidth;

            // 2) Data format
            var numberFormat = sourceColumn.Columns.NumberFormat;
            targetColumn.EntireColumn.NumberFormat = numberFormat;

            // 3) Centering
            targetColumn.EntireColumn.HorizontalAlignment = sourceColumn.HorizontalAlignment;
            targetColumn.EntireColumn.VerticalAlignment = sourceColumn.VerticalAlignment;
        }

        /// <summary>
        /// Formats sheet:
        /// #- centered headers
        /// #- bold headings w/ word wrap on
        /// #- auto-fit all columns
        /// #- "NULL" values grayed out
        /// #- top row frozen
        /// #- thick bottom border on top row
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        internal void Format(Worksheet worksheet)
        {
            Range originalSelection = worksheet.Application.Selection;
            lastColumn = worksheet.UsedRange.Columns.Count;
            lastRow = worksheet.UsedRange.Rows.Count;

            // If the user has selected the first row, we won't be free to modify it.
            MoveOffFirstRow(originalSelection);

            // Apply special formatting based on column names.
            FormatMRN(worksheet);
            FormatDates(worksheet);

            // Format the header row.
            FitHeader(worksheet);
            SetBorders(worksheet);
            FreezePane(worksheet);
            CenterlineTheMain(worksheet);
            WrapText(worksheet);
            MakeHeaderBold(worksheet);

            GrayOutTheNulls(worksheet);

            // Restore original selection.
            originalSelection.Select();
        }

        /// <summary>
        /// Formats all columns expand to fit their contents.
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        private void FitHeader(Worksheet worksheet)
        {
            Range firstCell = worksheet.Cells[1, 1];

            double dataColWidth;
            double desiredColWidth;
            double headerColWidth;
            int stringLength;

            // Run across the columns.
            for (int colOffset = 0; colOffset < lastColumn; colOffset++)
            {
                headerColWidth = 0.0;

                // Wrap this in a try/catch because modifying a column while user is editing a cell in it will throw an exception.
                try
                {
                    // Get width of HEADER cell...
                    stringLength = firstCell.Offset[0, colOffset].Value2.ToString().Length;
                    
                    // Compute width of HEADER cell.
                    headerColWidth = stringLength * POINTS_PER_CHAR;

                    // Get width of first DATA cell...
                    stringLength = firstCell.Offset[1, colOffset].Value2.ToString().Length;

                    // Compute width of DATA cell.
                    dataColWidth = stringLength * POINTS_PER_CHAR;

                    // Set the column width to the larger of the required HEADER cell width
                    // OR the required DATA cell width (but not wider than the MAX width).
                    desiredColWidth = Math.Max(headerColWidth * BOLD_BUMP, dataColWidth);
                    desiredColWidth = Math.Min(MAX_COLUMN_WIDTH, desiredColWidth);
                    firstCell.Offset[0, colOffset].Columns.ColumnWidth = desiredColWidth;
                }
                catch (System.Runtime.InteropServices.COMException) { }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // Then we've run out of data.
                    return;
                }
            }
        }

        /// <summary>
        /// Formats the column containing Medical Record Numbers to use 0####### numerical format.
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        private void FormatMRN(Worksheet worksheet)
        {
            Range firstCell = worksheet.Cells[1, 1];

            for (int colOffset = 0; colOffset < lastColumn; colOffset++)
            {
                try
                {
                    string columnName = firstCell.Offset[0, colOffset].Value2.ToString();

                    if (columnName.ToUpper().Contains("MRN"))
                    {
                        try
                        {
                            worksheet.Columns[colOffset + 1].NumberFormat = "0#######";
                        }
                        catch (System.Runtime.InteropServices.COMException) { }
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) 
                {
                    // Then we've run out of data.
                    return;
                }
            }
        }

        /// <summary>
        /// Formats the column containing dates to use "MM/DD/YYYY" format
        /// and datetimes to use "MM/DD/YYYY hh:mm:ss".
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        private void FormatDates(Worksheet worksheet)
        {
            Range firstCell = worksheet.Cells[1, 1];

            for (int colOffset = 0; colOffset < lastColumn; colOffset++)
            {
                Range topOfColumn = firstCell.Offset[0, colOffset];

                try 
                {
                    string columnName = topOfColumn.Value2.ToString();

                    if (columnName.ToLower().Contains("date") && Utilities.IsExcelDate(topOfColumn, lastRow))
                    {
                        try
                        {
                            worksheet.Columns[colOffset + 1].NumberFormat = "MM/DD/YYYY";
                        }
                        catch (System.Runtime.InteropServices.COMException) { }
                    }

                    if (columnName.ToLower().Contains("dttm") && Utilities.IsExcelDate(topOfColumn, lastRow))
                    {
                        try
                        {
                            worksheet.Columns[colOffset + 1].NumberFormat = "MM/DD/YYYY hh:mm:ss";
                        }
                        catch (System.Runtime.InteropServices.COMException) { }
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // Then we've run out of data.
                    return;
                }
            }
        }

        /// <summary>
        /// Freezes the top row so it's always visible as we scroll down.
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        private void FreezePane(Worksheet worksheet)
        {
            try
            {
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        /// <summary>
        /// Grays out all "NULL" values for better readability.
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        private void GrayOutTheNulls(Worksheet worksheet)
        {
            FormatCondition cond = worksheet.Cells.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, "NULL");
            cond.Font.Color = XlRgbColor.rgbLightGray;
        }

        /// <summary>
        /// Tests to see if user has selected the first row.
        /// </summary>
        /// <param name="selection">Range of selected region</param>
        /// <returns>bool</returns>
        private bool IsFirstRowSelected(Range selection)
        {
            int selectedRow = selection.Row;
            return selectedRow == 1;
        }

        /// <summary>
        /// Sets header text tp BP:D/
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        private void MakeHeaderBold(Worksheet worksheet)
        {
            Range topRow = worksheet.Cells[1, 1].EntireRow;
            topRow.Font.Bold = true;
        }

        /// <summary>
        /// Pushes the selection off the first row so we can modify first row.
        /// </summary>
        /// <param name="selection">Range of selected region</param>

        private void MoveOffFirstRow(Range selection)
        {
            if (IsFirstRowSelected(selection))
            {
                Worksheet sheet = selection.Worksheet;
                Range secondRow = sheet.Cells[2, 1];
                secondRow.Select();
            }
        }

        /// <summary>
        /// Formats desired row to have thick bottom border.
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        private void SetBorders(Worksheet worksheet)
        {
            Range topRow = worksheet.Cells[1, 1].EntireRow;

            // Trying to modify while user is editing a cell will result in an error.
            try
            {
                topRow.Borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;
                topRow.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                topRow.Borders[XlBordersIndex.xlEdgeBottom].Weight = 2;
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        /// <summary>
        /// Formats sheet to wrap text so long info or headers can be read in their entirety.
        /// </summary>
        /// <param name="worksheet">The active worksheet</param>

        private void WrapText(Worksheet worksheet)
        {
            // Trying to modify while user is editing a cell will result in an error.
            try
            {
                // Only apply to the header row.
                worksheet.Rows[1].Columns.WrapText = true;
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }
    }
}
