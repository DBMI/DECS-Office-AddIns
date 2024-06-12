using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using static System.Net.Mime.MediaTypeNames;
using System.Text.RegularExpressions;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    /**
    * @brief Class that extracts Time-In-Notes info from Epic Signal downloaded file.
    */
    internal class ExtractTime
    {
        private readonly string[] COLUMNSNEEDED = { "SER CID", "Provider Name", "Service Area", "Specialty", "Numerator", "Denominator", "Metric ID" };
        private readonly string[] PRIMARY_CARE_SPECIALTIES = { "Family Practice", "General Practice", "Medicine", "Pediatrics", "Sports Medicine"};
        private readonly string[] SURGICAL_SPECIALTY_PATTERNS = { "%SURG%"};
        private const string SPECIALTY_COLUMN_NAME = "Specialty";
        private const string METRIC_ID_COLUMN = "Metric ID";
        private const int TIME_IN_NOTES_METRIC_ID = 317;

        private Application application;
        private Workbook downloadedWorkbook;
        private Worksheet downloadedWorksheet;
        private Range metricsPageRange;
        private Worksheet metricsPageWorksheet;
        private string newFilename;
        private List<string> specialties;

        internal ExtractTime(Worksheet _worksheet)
        {
            application = Globals.ThisAddIn.Application;
            downloadedWorksheet = _worksheet;
            downloadedWorkbook = (Excel.Workbook)downloadedWorksheet.Parent;
            newFilename = GenerateNewFilename();

            // Create new worksheet for extracted numbers.
            ////TEMP////
            Dictionary<string, Worksheet> worksheets = Utilities.GetWorksheets();
            metricsPageWorksheet = worksheets["MetricData_extracted"];
            //metricsPageWorksheet = (Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            //metricsPageWorksheet.Name = downloadedWorksheet.Name + "_extracted";
            //// TEMP ////

            metricsPageRange = (Range)metricsPageWorksheet.Cells[1, 1];
        }

        private void CopyRow(Worksheet sourceWorksheet, Worksheet targetWorksheet, int sourceRowOffset = 0, int targetRowOffset = 0)
        {
            int rowNumber = sourceRowOffset + 1;
            application.StatusBar = "Copying row " + rowNumber.ToString();
            Utilities.CopyRow(sourceWorksheet, sourceRowOffset, targetWorksheet, targetRowOffset);
        }

        /// <summary>
        /// Deletes all but the columns we need.
        /// </summary>

        private void DeleteUnneededColumns()
        {
            Range sourceRange = (Range)downloadedWorksheet.Cells[1, 1];
            int lastColumnNumber = Utilities.FindLastCol(downloadedWorksheet);

            // Start deleting columns on the right edge & work left.
            for (int colOffset = lastColumnNumber - 1;  colOffset >= 0; colOffset--)
            {
                string columnName;

                try 
                {
                    columnName = sourceRange.Offset[0, colOffset].Value2.ToString();
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
                {
                    return;
                }

                if (!COLUMNSNEEDED.Contains(columnName))
                {
                    application.StatusBar = "Deleting column '" + columnName + "'";

                    try
                    {
                        downloadedWorksheet.Columns[colOffset + 1].Delete();
                    }
                    catch (Exception ex)
                    {
                        application.StatusBar = false;
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Deletes all but the rows we need.
        /// </summary>

        private void DeleteUnneededRows()
        {
            List<string> metricIdColumnName = new List<string> { METRIC_ID_COLUMN };
            Dictionary<string, Range> columnDict = Utilities.GetColumnRangeDictionary(downloadedWorksheet, metricIdColumnName);
            Range metricIdColumn = columnDict[METRIC_ID_COLUMN];

            int numRowsDeleted = 0;
            int lastRow = Utilities.FindLastRow(downloadedWorksheet);
            int rowNum = 2;
            int rowOffset;

            while (rowNum < lastRow - numRowsDeleted)
            {
                // Allow for header row.
                rowOffset = rowNum - 1;

                string value = metricIdColumn.Offset[rowOffset, 0].Value2.ToString();

                if (int.TryParse(value, out int metricValue))
                {
                    if (metricValue == TIME_IN_NOTES_METRIC_ID)
                    { 
                        rowNum++;
                    }
                    else
                    {
                        // Delete this row.
                        application.StatusBar = "Deleting row " + rowNum.ToString();
                        downloadedWorksheet.Rows[rowNum].Delete();
                        numRowsDeleted++;
                    }
                }
            }
        }

        internal void Extract()
        {
            //DeleteUnneededColumns();
            //downloadedWorkbook.SaveAs(newFilename);

            // Copy over the header to the extracted metrics sheet.
            //CopyRow(sourceWorksheet: downloadedWorksheet, targetWorksheet: metricsPageWorksheet);

            // Populate new sheet with just Time-In-Note rows.
            //ExtractMetricsToNewSheet();

            // Delete the main sheet & save the file.
            //downloadedWorksheet.Delete();
            //application.StatusBar = "Saving workbook...";
            //downloadedWorkbook.Save();

            GetSpecialties();

            // Build new sheet for each specialty.
            ExtractMetricsBySpecialty();

            application.StatusBar = false;
        }

        private void ExtractMetricsBySpecialty()
        {
            // From where will we read the specialty on the extracted metrics sheet?
            Dictionary<string, Range> columns = Utilities.GetColumnRangeDictionary(metricsPageWorksheet);
            Range specialtyColumn = columns[SPECIALTY_COLUMN_NAME];

            if (specialtyColumn == null)
            {
                return;
            }

            int lastRowNum = Utilities.FindLastRow(metricsPageWorksheet);

            foreach (string specialtyName in specialties)
            {
                // Create new worksheet for this department.
                Worksheet specialtyWorksheet = (Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                specialtyWorksheet.Name = specialtyName;
                int specialtyRowOffset = 1;
                application.StatusBar = "Extracting specialty: " + specialtyName;

                // Run down the rows.
                for (int metricsPageRowOffset = 1; metricsPageRowOffset < lastRowNum; metricsPageRowOffset++)
                {
                    string thisSpecialty;

                    try 
                    {
                        thisSpecialty = specialtyColumn.Offset[metricsPageRowOffset, 0].Value2.ToString();
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex) 
                    {
                        return;
                    }                    

                    if (thisSpecialty == specialtyName)
                    {
                        CopyRow(sourceWorksheet: metricsPageWorksheet,
                                targetWorksheet: specialtyWorksheet,
                                sourceRowOffset: metricsPageRowOffset,
                                targetRowOffset: specialtyRowOffset);
                        specialtyRowOffset++;
                    }
                }
            }
        }

        private void ExtractMetricsToNewSheet()
        {
            List<string> metricIdColumnName = new List<string> { METRIC_ID_COLUMN };
            Dictionary<string, Range> columnDict = Utilities.GetColumnRangeDictionary(downloadedWorksheet, metricIdColumnName);
            Range metricIdColumn = columnDict[METRIC_ID_COLUMN];

            int lastRowNum = Utilities.FindLastRow(downloadedWorksheet);
            int metricsPageRowOffset = 1;

            // Skip the header.
            for (int sourceRowOffset = 1; sourceRowOffset < lastRowNum; sourceRowOffset++)
            {
                string value = metricIdColumn.Offset[sourceRowOffset, 0].Value2.ToString();

                if (int.TryParse(value, out int metricValue))
                {
                    if (metricValue == TIME_IN_NOTES_METRIC_ID)
                    {
                        CopyRow(sourceWorksheet: downloadedWorksheet,
                                targetWorksheet: metricsPageWorksheet, 
                                sourceRowOffset: sourceRowOffset, 
                                targetRowOffset: metricsPageRowOffset);
                        metricsPageRowOffset++;
                    }
                }
            }
        }

        private string GenerateNewFilename()
        {
            // Save a copy of the revised workbook.
            string filename = downloadedWorkbook.FullName;
            string directory = System.IO.Path.GetDirectoryName(filename);
            string justTheFilename = System.IO.Path.GetFileNameWithoutExtension(filename);
            string newFilename = System.IO.Path.Combine(
                directory,
                justTheFilename + ".xlsx"
            );

            return newFilename;
        }

        private void GetSpecialties()
        {
            specialties = new List<string>();
            Dictionary<string, Range> columns = Utilities.GetColumnRangeDictionary(metricsPageWorksheet);
            Range specialtyColumn = columns[SPECIALTY_COLUMN_NAME];

            if (specialtyColumn == null)
            {
                return;
            }

            int specialtyColumnOffset = specialtyColumn.Column - 1;
            int lastRow = Utilities.FindLastRow(metricsPageWorksheet);

            GetSpecialtiesByName(specialtyColumnOffset, lastRow);
            GetSpecialtiesByPattern(specialtyColumnOffset, lastRow);

            specialties.Sort();
        }

        private void GetSpecialtiesByName(int specialtyColumnOffset, int lastRow)
        {
            application.StatusBar = "Reading specialties by name....";
            string specialtyName;

            for (int rowOffset = 1; rowOffset < lastRow; rowOffset++)
            {
                try
                {
                    specialtyName = metricsPageRange.Offset[rowOffset, specialtyColumnOffset].Value2.ToString();

                    if (PRIMARY_CARE_SPECIALTIES.Contains(specialtyName) && !specialties.Contains(specialtyName))
                    {
                        specialties.Add(specialtyName);
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
                {
                    return;
                }

                application.StatusBar = "Reading specialties by name: row " + rowOffset;
            }
        }

        private void GetSpecialtiesByPattern(int specialtyColumnOffset, int lastRow)
        {
            string specialtyName;

            foreach (string pattern in SURGICAL_SPECIALTY_PATTERNS)
            {
                Regex regex = new Regex(pattern);

                for (int rowOffset = 1; rowOffset < lastRow; rowOffset++)
                {
                    try
                    {
                        specialtyName = metricsPageRange.Offset[rowOffset, specialtyColumnOffset].Value2.ToString();
                        Match line_match = regex.Match(specialtyName);

                        if (line_match.Success && !specialties.Contains(specialtyName))
                        {
                            specialties.Add(specialtyName);
                        }
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
                    {
                        return;
                    }

                    application.StatusBar = "Reading specialties by pattern: row " + rowOffset;
                }
            }
        }
    }
}
