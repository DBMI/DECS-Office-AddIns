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
    internal class Formatter
    {
        internal Formatter() { }

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
        private void FitHeader(Range row)
        {
            // Trying to modify while user is editing a cell will result in an error.
            try
            {
                row.AutoFit();
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }
        private void FreezePane(Worksheet sheet)
        {
            try
            {
                sheet.Application.ActiveWindow.SplitRow = 1;
                sheet.Application.ActiveWindow.FreezePanes = true;
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        private void GrayOutTheNulls(Worksheet worksheet)
        {
            FormatCondition cond = worksheet.Cells.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, "NULL");
            cond.Font.Color = XlRgbColor.rgbLightGray;
        }

        private bool IsFirstRowSelected(Range selection)
        {
            int selectedRow = selection.Row;
            return selectedRow == 1;
        }
        
        private void MoveOffFirstRow(Range selection)
        {
            if (IsFirstRowSelected(selection))
            {
                Worksheet sheet = selection.Worksheet;
                Range secondRow = sheet.Cells[2, 1];
                secondRow.Select();
            }
        }

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
