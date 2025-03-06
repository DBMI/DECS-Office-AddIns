using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    internal class Timecard
    {
        private Range lastMonthCumulativeHours;
        private Range thisMonthCumulativeHours;
        private DateTime newFileDate;
        private Range thisMonthNewHours;
        private Workbook thisWorkbook;
        private Range topOfNewHours;

        internal Timecard() { }

        private void BuildGlobals(Worksheet worksheet)
        {
            thisWorkbook = (Workbook)worksheet.Parent;
        }

        private Worksheet CopyToNewSheet(Worksheet lastMonthSheet)
        {
            lastMonthSheet.Copy(Type.Missing, thisWorkbook.Sheets[thisWorkbook.Sheets.Count]);   // copy
            string newSheetName = newFileDate.ToString("MMMM yyyy");
            Worksheet thisMonthSheet = thisWorkbook.Sheets[thisWorkbook.Sheets.Count];
            thisMonthSheet.Name = newSheetName;                                                 // rename
            return thisMonthSheet;
        }

        internal void Extend(Worksheet worksheet)
        {
            // Determine some global values.
            BuildGlobals(worksheet);

            if (!SaveNextMonthVersion()) { return; }

            // Find the latest sheet.
            Worksheet lastMonthSheet = Utilities.FindLastWorksheet(thisWorkbook);

            if (lastMonthSheet == null) { return; }

            // Point to its first cell in column "K".
            lastMonthCumulativeHours = (Range)lastMonthSheet.Cells[2, 11];

            // Copy all to a new sheet with next month name.
            Worksheet thisMonthSheet = CopyToNewSheet(lastMonthSheet);

            // Point to ITS first cell in column "J".
            thisMonthNewHours = (Range)thisMonthSheet.Cells[2, 10];

            // Remember this one for the missing hours formula.
            topOfNewHours = (Range)thisMonthSheet.Cells[2, 10];

            // And first cell in column "K".
            thisMonthCumulativeHours = (Range)thisMonthSheet.Cells[2, 11];

            // Zero out the values in the new sheet's "Actual hours this month" column.
            ZeroOutNewMonthActualHours();

            // Shift all the formulas in the new sheet.
            UpdateHoursFormulas();

            // Shift the # days formula in the new sheet.
            UpdateMissingHoursFormulas();

            // Save revised workbook.
            thisWorkbook.Save();
        }

        private bool PastLastRow(string cellContents)
        {
            return !cellContents.Contains("=") && !cellContents.Contains("see next row");
        }

        private string NewFilename()
        {
            string filename = thisWorkbook.FullName;
            string directory = System.IO.Path.GetDirectoryName(filename);
            string justTheFilename = System.IO.Path.GetFileNameWithoutExtension(filename);

            // Parse year, month from string like "DFMResearchProjects_Kevin_2024_11.xlsx".
            Regex regex = new Regex(@"(?<preamble>\D+)(_|\s)(?<year>\d{4})(_|\s)(?<month>\d{1,2})$");
            Match match = regex.Match(justTheFilename);

            if (match.Success)
            {
                if (int.TryParse(match.Groups["year"].Value, out int year) &&
                    int.TryParse(match.Groups["month"].Value, out int month))
                {
                    DateTime oldFileDate = new DateTime(year, month, 1);
                    newFileDate = oldFileDate.AddMonths(1);

                    string newFilename = System.IO.Path.Combine(
                    directory,
                    match.Groups["preamble"].Value + "_" +
                    newFileDate.ToString("yyyy_MM") + ".xlsx");

                    return newFilename;
                }
            }

            return string.Empty;
        }

        private bool SaveNextMonthVersion()
        {
            bool success = false;

            // Generate new name.
            string newFilename = NewFilename();

            if (!string.IsNullOrEmpty(newFilename))
            {
                thisWorkbook.SaveAs(newFilename);
                success = true;
            }

            return success;
        }

        private void UpdateHoursFormulas()
        {
            bool done = false;
            string newFormula;

            while (!done) 
            {
                newFormula = "='" +
                            lastMonthCumulativeHours.Worksheet.Name + "'!" +
                            lastMonthCumulativeHours.Address + " + '" +
                            thisMonthNewHours.Worksheet.Name + "'!" +
                            thisMonthNewHours.Address;
                thisMonthCumulativeHours.Formula = newFormula;

                // Bump down to the next row.
                lastMonthCumulativeHours = lastMonthCumulativeHours.Offset[1, 0];
                thisMonthCumulativeHours = thisMonthCumulativeHours.Offset[1, 0];
                thisMonthNewHours = thisMonthNewHours.Offset[1, 0];

                // Have we passed the last formula?
                done = PastLastRow(thisMonthCumulativeHours.Formula);
            }
        }

        private void UpdateMissingHoursFormulas()
        {
            string newFormula;

            // Insert formula to sum all the new hours entries to compute hours worked so far this month.
            Range lastCellInNewHoursColumn = thisMonthNewHours.Offset[-1, 0];
            newFormula = "=SUM(" + topOfNewHours.Address + ":" + lastCellInNewHoursColumn.Address + ")";
            thisMonthNewHours.Formula = newFormula;

            // Insert formula to compute the available work hours so far this month.
            thisMonthNewHours = thisMonthNewHours.Offset[1, 0];
            newFormula = "= 8* (NETWORKDAYS(DATE(" +
                        newFileDate.Year.ToString() + ", " +
                        newFileDate.Month.ToString() + ", 1), TODAY()) - 1)";
            thisMonthNewHours.Formula = newFormula;

            // Compute uncharged hours so far this month.
            thisMonthNewHours = thisMonthNewHours.Offset[1, 0];
            newFormula = "=" + thisMonthNewHours.Offset[-1, 0].Address + "-" +
                                        thisMonthNewHours.Offset[-2, 0].Address;
            thisMonthNewHours.Formula = newFormula;
        }

        private void ZeroOutNewMonthActualHours()
        {
            bool done = false;
            int rowOffset = 0;

            while (!done)
            {
                thisMonthNewHours.Offset[rowOffset, 0].Value = 0;

                // Bump down to the next row.
                rowOffset++;

                // Have we passed the last formula?
                done = PastLastRow(thisMonthCumulativeHours.Offset[rowOffset, 0].Formula);
            }
        }
    }
}