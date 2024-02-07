using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Class that formats Excel output to look "nice".
     */
    internal class Formatter
    {
        internal Formatter() { }

        /// <summary>
        /// Formats sheet:
        /// #- centered columns
        /// #- bold headings w/ word wrap on
        /// #- auto-fit all columns
        /// #- "NULL" values grayed out
        /// #- top row frozen
        /// #- thick bottom border on top row
        /// </summary>
        /// <param name="sheet">The active worksheet</param>
        
        internal void Format(Worksheet worksheet)
        {
            Range originalSelection = worksheet.Application.Selection;

            // If the user has selected the first row, we won't be free to modify it.
            MoveOffFirstRow(originalSelection);

            CenterlineTheMain(worksheet);
            WrapText(worksheet);
            FreezePane(worksheet);
            GrayOutTheNulls(worksheet);

            Range topRow = worksheet.Cells[1, 1].EntireRow;            
            topRow.Font.Bold = true;
            FitHeader(topRow);
            SetBorders(topRow);

            // Restore original selection.
            originalSelection.Select();
        }

        /// <summary>
        /// Formats all columns to be centered horizontally & vertically.
        /// </summary>
        /// <param name="sheet">The active worksheet</param>
        
        private void CenterlineTheMain(Worksheet sheet)
        {
            Workbook workbook = sheet.Application.ActiveWorkbook;

            // Trying to modify while user is editing a cell will result in an error.
            try
            {
                Style style = workbook.Styles.Add("CenteredHeadings");
                style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = XlVAlign.xlVAlignCenter;
                sheet.Columns.Style = "CenteredHeadings";
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        /// <summary>
        /// Formats all columns expand to fit their contents.
        /// </summary>
        /// <param name="row">Range of desired row--usually the top row.</param>
        
        private void FitHeader(Range row)
        {
            // Trying to modify while user is editing a cell will result in an error.
            try
            {
                row.AutoFit();
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        /// <summary>
        /// Freezes the top row so it's always visible as we scroll down.
        /// </summary>
        /// <param name="sheet">The active worksheet</param>
        
        private void FreezePane(Worksheet sheet)
        {
            try
            {
                sheet.Application.ActiveWindow.SplitRow = 1;
                sheet.Application.ActiveWindow.FreezePanes = true;
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        /// <summary>
        /// Grays out all "NULL" values for better readability.
        /// </summary>
        /// <param name="sheet">The active worksheet</param>
        
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
        /// <param name="row">Range of desired row--usually the top row.</param>
        
        private void SetBorders(Range row)
        {

            // Trying to modify while user is editing a cell will result in an error.
            try
            {
                row.Borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;
                row.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                row.Borders[XlBordersIndex.xlEdgeBottom].Weight = 2;
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        /// <summary>
        /// Formats sheet to wrap text so long info or headers can be read in their entirety.
        /// </summary>
        /// <param name="sheet">The active worksheet</param>
        
        private void WrapText(Worksheet sheet)
        {
            // Trying to modify while user is editing a cell will result in an error.
            try
            {
                sheet.Columns.WrapText = true;
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }
    }
}
