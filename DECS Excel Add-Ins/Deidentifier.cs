using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.UI;
using System.Windows.Forms;
using System.Xml.Linq;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    internal class Deidentifier
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private const string dateOnlyPattern = @"\d{1,2}\/\d{1,2}\/\d{4}[\s\.](?!\d)";
        private Regex dateOnlyRegex;
        private const string dateTimePattern = @"\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s[AP]M";
        private Regex dateTimeRegex;
        private const string monthDayOnlyPattern = @"\d{1,2}\/\d{1,2}(?![\/\d])";
        private Regex monthDayOnlyRegex;
        private const string monthSpelledOutPattern = @"\w{4,8}\s*\d{4}";
        private Regex monthSpelledOutRegex;
        private int dayOffset;
        private int monthOffset;
        private int lastRow;
        private Range selectedColumnRng;
        private List<Range> selectedColumnsRng;
        private TimeSpan deltaT;
        private byte[] tmpSource;
        private byte[] tmpHash;

        internal Deidentifier()
        {
            application = Globals.ThisAddIn.Application;
        }

        // https://learn.microsoft.com/en-us/troubleshoot/developer/visualstudio/csharp/language-compilers/compute-hash-values
        private static string ByteArrayToString(byte[] arrInput)
        {
            int i;
            StringBuilder sOutput = new StringBuilder(arrInput.Length);

            for (i = 0; i < arrInput.Length; i++)
            {
                sOutput.Append(arrInput[i].ToString("X2"));
            }

            return sOutput.ToString();
        }

        private string TweakDateOnly(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddDays(dayOffset);
            string convertedDateString = payloadTweaked.ToString("M/d/yyyy");

            // Special case: did the Regex absorb a trailing period or space?
            if (dateString.EndsWith("."))
            {
                convertedDateString += ".";
            }
            
            if (dateString.EndsWith(" "))
            {
                convertedDateString += " ";
            }

            return convertedDateString;
        }

        private string TweakDateTime(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddDays(dayOffset) + deltaT;
            return payloadTweaked.ToString("M/d/yyyy h:mm tt");
        }

        private string TweakMonthDay(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddDays(dayOffset);
            return payloadTweaked.ToString("M/d");
        }

        private string TweakMonthSpelledOut(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddMonths(monthOffset);
            return payloadTweaked.ToString("MMMM yyyy");
        }

        private bool FindSelectedColumn(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedColumnRng = Utilities.GetSelectedCol(application, lastRow);

            if (selectedColumnRng is null)
            {
                // Then ask user to select one column.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames, MultiSelect: false))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        selectedColumnRng = Utilities.TopOfNamedColumn(worksheet, form.selectedColumns[0]);
                        success = true;
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Then we're done here.
                        return success;
                    }
                }
            }
            else
            {
                success = true;
            }

            return success;
        }

        private bool FindSelectedColumns(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedColumnsRng = Utilities.GetSelectedCols(application, lastRow);

            if (selectedColumnsRng is null)
            {
                // Then ask user to select columns of interest.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames, MultiSelect: true))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        foreach (string selectedColumnName in form.selectedColumns)
                        {
                            Range thisRng = Utilities.TopOfNamedColumn(worksheet, selectedColumnName);
                            selectedColumnsRng.Add(thisRng);
                        }

                        success = true;
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Then we're done here.
                        return success;
                    }
                }
            }
            else
            {
                success = true;
            }

            return success;
        }

        internal void GenerateHash(Worksheet worksheet)
        {
            lastRow = worksheet.UsedRange.Rows.Count;

            if (FindSelectedColumns(worksheet))
            {
                // Make room for new column.
                Range hashColumn = Utilities.InsertNewColumn(range: selectedColumnsRng.Last(), newColumnName: "Scrambled Identifier", side: InsertSide.Right);

                string sourceData;
                Range target;
                byte[] tmpHash;
                byte[] tmpSource;

                for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
                {
                    target = (Range)worksheet.Cells[rowNumber, hashColumn.Column];
                    sourceData = Utilities.CombineColumns(worksheet, rowNumber, selectedColumnsRng);

                    if (!string.IsNullOrEmpty(sourceData))
                    {
                        // Create a byte array from source data.
                        tmpSource = ASCIIEncoding.ASCII.GetBytes(sourceData);

                        // Initialize a SHA256 hash object.
                        using (SHA256 mySHA256 = SHA256.Create())
                        {
                            tmpHash = mySHA256.ComputeHash(tmpSource);
                            target.Value = ByteArrayToString(tmpHash);
                        }
                    }
                }
            }
        }

        internal void ObscureDateTime(Worksheet worksheet)
        {
            // Instantiate random number generator and random day, time offsets.
            Random rnd = new Random();
            dayOffset = rnd.Next(-7, 7);
            monthOffset = rnd.Next(-2, 2);
            int hourOffset = rnd.Next(-3, 3);
            int minuteOffset = rnd.Next(-20, 20);
            deltaT = new TimeSpan(hourOffset, minuteOffset, 0);

            // Instantiate reusable Regexes.
            dateOnlyRegex = new Regex(dateOnlyPattern);
            dateTimeRegex = new Regex(dateTimePattern);
            monthDayOnlyRegex = new Regex(monthDayOnlyPattern);
            monthSpelledOutRegex = new Regex(monthSpelledOutPattern);

            lastRow = worksheet.UsedRange.Rows.Count;

            if (FindSelectedColumn(worksheet))
            {
                string selectedColumnName = selectedColumnRng.Value.ToString();
                string newColumnName = selectedColumnName + " (Date/Time Altered)";

                // Make room for new column.
                Range ditheredColumn = Utilities.InsertNewColumn(range: selectedColumnRng, 
                                                                 newColumnName: newColumnName, 
                                                                 side: InsertSide.Right);

                string sourceData;
                Range target;

                for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
                {
                    target = (Range)worksheet.Cells[rowNumber, ditheredColumn.Column];
                    sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value;

                    // Modify & stuff into target cell.
                    if (!string.IsNullOrEmpty(sourceData))
                    {
                        sourceData = ProcessOneRule(sourceData, dateTimeRegex, TweakDateTime);
                        sourceData = ProcessOneRule(sourceData, dateOnlyRegex, TweakDateOnly);
                        sourceData = ProcessOneRule(sourceData, monthDayOnlyRegex, TweakMonthDay);
                        sourceData = ProcessOneRule(sourceData, monthSpelledOutRegex, TweakMonthSpelledOut);
                    }

                    target.Value = sourceData;
                }
            }
        }

        private string ProcessOneRule(string sourceData, Regex regex, Func<string, string> convert)
        {
            Match match = regex.Match(sourceData);
            string targetData = string.Empty;

            while (match.Success)
            {
                string beforeMatch = sourceData.Substring(0, match.Index);
                targetData += beforeMatch;
                targetData += convert(match.Value.ToString());

                // Trim to just what's AFTER the match.
                sourceData = sourceData.Substring(match.Index + match.Value.Length);
                match = regex.Match(sourceData);
            }

            // Append whatever's left over.
            targetData += sourceData;

            return targetData;
        }
    }
}