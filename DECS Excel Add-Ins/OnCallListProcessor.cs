using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief One assignment.
     */
    internal class OnCallAssignment
    {
        private string name;
        private DateTime start;
        private DateTime end;

        internal OnCallAssignment(string name, DateTime start, DateTime end)
        {
            this.name = name;
            this.start = start;
            this.end = end;
        }

        // If we've already turned the string into a DateRange object.
        internal OnCallAssignment(string name, DateRange dateRange)
        {
            this.name = name;
            this.start = dateRange.Start();
            this.end = dateRange.End();
        }

        // If we need to modify the DateRange using info in the name string.
        internal OnCallAssignment(string nameContent, DateRange dateRange, int index)
        {
            DateTime nominalStart = dateRange.Start();
            DateTime nominalEnd = dateRange.End();

            // Separate the name and date portions in the name (like "Alice(29-31)").
            Regex regex = new Regex(@"(?<alpha>\w+)\s*\((?<start_day>\d{1,2})-(?<end_day>\d{1,2})\)");
            
            Match match = regex.Match(nameContent);

            if (match.Success)
            {
                this.name = match.Groups["alpha"].Value;

                if (int.TryParse(match.Groups["start_day"].Value, out int start_day) &&
                    int.TryParse(match.Groups["end_day"].Value, out int end_day))
                {
                    if (index == 0)
                    {
                        // Use the extracted days to modify the START date of the date range.
                        this.start = new DateTime(nominalStart.Year,
                            nominalStart.Month, start_day);
                        this.end = new DateTime(nominalStart.Year,
                            nominalStart.Month, end_day);
                    }
                    else if (index == 1)
                    {
                        // Use the extracted days to modify the END date of the date range.
                        this.start = new DateTime(nominalEnd.Year,
                            nominalEnd.Month, start_day);
                        this.end = new DateTime(nominalEnd.Year,
                            nominalEnd.Month, end_day);
                    }
                    else
                    {
                        throw new ArgumentOutOfRangeException($"Expected 'index' to be 0 or 1, not {index}.");
                    };
                }
            }
        }

        internal string Output()
        {
            return ("('" + start.ToString("yyyy-MM-dd") + "', '" + end.ToString("yyyy-MM-dd") + "', '" + name + "')");
        }
    }
    /**
     * @brief Builds SQL-friendly file with name of on-call person and the start/end dates for their assignment.
     */
    internal class OnCallListProcessor
    {
        private readonly string[] DATECOLUMNNAMES = { "DATES" };
        private readonly string[] NAMECOLUMNNAMES = { "Call" };

        private Range datesColumnRng;
        private Range namesColumnRng;

        private const string HEADER = "INSERT INTO #ON_CALL_LIST (ON_CALL_START_DATE, ON_CALL_END_DATE, PHYSICIAN)\r\nVALUES\r\n";
        private const string PREAMBLE = "USE [REL_CLARITY];\r\n\r\n";

        internal OnCallListProcessor()
        {
        }

        private bool FindRequiredColumnsByName(Worksheet worksheet)
        {
            bool foundDateColumn = false;
            bool foundNameColumn = false;

            // Can we find the dates, names columns by their header names?
            Dictionary<string, Range> columnsDict = Utilities.GetColumnRangeDictionary(worksheet);

            if (columnsDict != null)
            {
                foreach (string dateColumName in DATECOLUMNNAMES)
                {
                    try
                    {
                        datesColumnRng = columnsDict[dateColumName];
                        foundDateColumn = true;
                        break;      // Stop searching
                    }
                    catch (System.Collections.Generic.KeyNotFoundException) { }
                }

                foreach (string nameColumName in NAMECOLUMNNAMES)
                {
                    try
                    {
                        namesColumnRng = columnsDict[nameColumName];
                        foundNameColumn = true;
                        break;      // Stop searching
                    }
                    catch (System.Collections.Generic.KeyNotFoundException) { }
                }
            }

            return foundDateColumn & foundNameColumn;
        }

        private bool FindRequiredColumns(Worksheet worksheet)
        {
            bool success = false;

            if (!FindRequiredColumnsByName(worksheet))
            {
                // Then ask user to select the columns.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseOnCallColumnsForm form = new ChooseOnCallColumnsForm(columnNames))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        string selectedDateColumnName = form.selectedDateColumn;
                        datesColumnRng = Utilities.TopOfNamedColumn(worksheet, selectedDateColumnName);
                        string selectedNameColumnName = form.selectedNameColumn;
                        namesColumnRng = Utilities.TopOfNamedColumn(worksheet, selectedNameColumnName);
                        success = true;
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Then we're done here.
                        success = false;
                    }
                }
            }
            else
            {
                success = true;
            }

            return success;
        }

        private void Output(Worksheet worksheet, List<OnCallAssignment> assignments)
        {
            // Initialize the output .SQL file.
            Workbook workbook = worksheet.Parent;
            string workbookFilename = workbook.FullName;

            (StreamWriter writer, string outputFilename) = Utilities.OpenOutput(
                inputFilename: workbookFilename,
                filenameAddon: "_list",
                filetype: ".sql"
            );

            writer.Write(PREAMBLE);
            writer.Write(HEADER);

            int assignmentsProcessed = 0;

            foreach (OnCallAssignment assignment in assignments)
            {
                writer.Write(assignment.Output());
                assignmentsProcessed++;
                string line_ending = ",\r\n";

                if (assignmentsProcessed == assignments.Count)
                {
                    line_ending = ";\r\n";
                }

                writer.Write(line_ending);
            }

            writer.Close();

            // Show the resulting file.
            Process.Start(outputFilename);
        }

        internal void Scan(Worksheet worksheet)
        {
            List<OnCallAssignment> assignments = new List<OnCallAssignment>();

            if (FindRequiredColumns(worksheet))
            {
                int lastRowInSheet = worksheet.UsedRange.Rows.Count;
                int assumedYear = DateTime.Now.Year;
                DateRange previousDateRange = null;

                for (int rowOffset = 1; rowOffset < lastRowInSheet; rowOffset++)
                {
                    try 
                    {
                        string dateContent = datesColumnRng.Offset[rowOffset, 0].Value2.ToString();
                        string nameContent = namesColumnRng.Offset[rowOffset, 0].Value2.ToString();
                        DateRange dateRange = new DateRange(dateContent, assumedYear);

                        // Check for shift in year.
                        //  (Like when last row was 12/27/24-12/31/24 and next row is "1/1-1/2", which
                        //   automatically gets interpreted as 1/1/24-1/2/24.)

                        if (previousDateRange != null && dateRange.Start() < previousDateRange.Start()) 
                        {
                            dateRange.AddYear();
                            assumedYear++;
                        }

                        previousDateRange = dateRange;

                        if (dateRange.Valid())
                        {
                            // Does the name field contain split assignments like "Jones(29-31)/Smith(1-4)"?
                            if (nameContent.Any(char.IsDigit))
                            {
                                string[] namePieces = nameContent.Split('/');
                                int pieceIndex = 0;

                                foreach (string piece in namePieces)
                                { 
                                    OnCallAssignment assignment = new OnCallAssignment(piece, dateRange, pieceIndex);
                                    assignments.Add(assignment);
                                    pieceIndex++;
                                }
                            }
                            else
                            {
                                assignments.Add(new OnCallAssignment(nameContent, dateRange));
                            }
                        }
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                }

                Output(worksheet, assignments);
            }
        }
    }
}
