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
//        private const string dateTimePattern = @"(?<month>\d{1,2})\/(?<day>\d{1,2})\/(?<year>\d{4})\s+(?<hour>\d{1,2}):(?<minute>\d{2})\s(?<ampm>(A|P)M)";
        private const string dateTimePattern = @"\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s[AP]M";
        private int dayOffset;
        private int lastRow;
        private Random rnd;
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

        private string ConvertDateTime(Match match)
        {
            string convertedDateTimeString = string.Empty;

            DateTime payload = DateTime.ParseExact(match.Value, "M/d/yyyy h:mm tt", CultureInfo.InvariantCulture);
            DateTime payloadConverted = payload.AddDays(dayOffset);
            payloadConverted += deltaT;
            return payloadConverted.ToString("M/d/yyyy h:mm tt");
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
            // Instantiate reusable Regex.
            Regex dateTimeRegex = new Regex(dateTimePattern);
            
            // Instantiate random number generator and random day, time offsets.
            rnd = new Random();
            dayOffset = rnd.Next(-7, 7);
            int hourOffset = rnd.Next(-3, 3);
            int minuteOffset = rnd.Next(-20, 20);
            deltaT = new TimeSpan(hourOffset, minuteOffset, 0);

            lastRow = worksheet.UsedRange.Rows.Count;

            if (FindSelectedColumn(worksheet))
            {                
                // Make room for new column.
                Range ditheredColumn = Utilities.InsertNewColumn(range: selectedColumnRng, newColumnName: "Date/Time Altered", side: InsertSide.Right);

                string sourceData;
                Range target;

                for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
                {
                    target = (Range)worksheet.Cells[rowNumber, ditheredColumn.Column];
                    sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value;
                    string targetData = string.Empty;

                    if (!string.IsNullOrEmpty(sourceData))
                    {
                        Match match = dateTimeRegex.Match(sourceData);

                        while (match.Success)
                        {
                            string beforeMatch = sourceData.Substring(0, match.Index);
                            targetData += beforeMatch;
                            targetData += ConvertDateTime(match);

                            // Trim to just what's AFTER the match.
                            sourceData = sourceData.Substring(match.Index + match.Value.Length);
                            match = dateTimeRegex.Match(sourceData);
                        }

                        target.Value = targetData + sourceData;
                    }
                }
            }
        }
    }
}