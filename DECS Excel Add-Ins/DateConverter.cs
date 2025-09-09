using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Centralized class for all date conversion methods.
     */
    internal class DateConverter
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private List<Range> selectedColumnsRng;
        private IDictionary<string, string> supportedDateFormats;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        internal DateConverter()
        {
            application = Globals.ThisAddIn.Application;

            supportedDateFormats = new Dictionary<string, string>();
            supportedDateFormats.Add("MM/dd/yyyy", "(\\d{1,2}\\/\\d{1,2}\\/\\d{4})");
            supportedDateFormats.Add("MM-dd-yyyy", "(\\d{1,2}-\\d{1,2}-\\d{4})");
            supportedDateFormats.Add("dd MMMM yyyy", "(\\d{1,2} \\w{3,9}\\.? \\d{4})");
            supportedDateFormats.Add("MMMM dd yyyy", "([a-zA-Z]{3,9}\\.? \\d{1,2},? \\d{4})");
            supportedDateFormats.Add("MMMM dd", "([a-zA-Z]\\.? \\d{1,2})");
        }

        /// <summary>
        /// Convert all dates found in the string to the desired format.
        /// </summary>
        /// <param name="note">Long string of patient notes.</param>
        /// <param name="desiredFormat">Desired date format</param>
        /// <returns>string</returns>
        internal string Convert(string note, string desiredFormat)
        {
            foreach (KeyValuePair<string, string> entry in supportedDateFormats)
            {
                if (entry.Key == desiredFormat)
                    continue;

                foreach (Match match in Regex.Matches(note, entry.Value))
                {
                    if (match.Success)
                    {
                        string dateString = match.Value.ToString();
                        log.Debug("Rule matched: " + dateString);

                        if (DateTime.TryParse(dateString, out DateTime dateValue))
                        {
                            string dateConverted = dateValue.ToString(desiredFormat);
                            log.Debug("Converted '" + dateString + "' to '" + dateConverted + "'.");
                            note = note.Replace(dateString, dateConverted);
                        }
                    }
                }
            }

            return note;
        }

        private bool FindSelectedColumns(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedColumnsRng = Utilities.GetSelectedCols(application);

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

        /// <summary>
        /// Provides all the keys from the supportedDateFormats Dictionary, for use in pull-down box.
        /// </summary>
        /// <returns>List<string></returns>
        internal List<string> SupportedDateFormats()
        {
            return new List<string>(supportedDateFormats.Keys);
        }

        /// <summary>
        /// Adds new column with string version of dates in selected column.
        /// </summary>
        /// <param name="note">Long string of patient notes.</param>
        /// <param name="desiredFormat">Desired date format</param>
        /// <returns>string</returns>
        internal void ToText(Worksheet worksheet)
        {
            if (FindSelectedColumns(worksheet))
            {
                foreach (Range selectedColumnRng in selectedColumnsRng)
                {
                    ToText(selectedColumnRng);
                }
            }
        }

        internal void ToText(Range selectedColumnRng)
        {
            string selectedColumnName = selectedColumnRng.Value.ToString();
            string newColumnName = selectedColumnName + " Text";

            // Make room for new column.
            Range newColumn = Utilities.InsertNewColumn(range: selectedColumnRng,
                                                             newColumnName: newColumnName,
                                                             side: InsertSide.Right);
            newColumn.NumberFormat = "@";

            DateTime sourceData;
            Range target;
            Worksheet worksheet = selectedColumnRng.Worksheet;
            int rowNumber = 1;
            int numConsecutiveFailures = 0;

            while (true)
            {
                rowNumber++;
                target = (Range)worksheet.Cells[rowNumber, newColumn.Column];

                try
                {
                    sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value;
                    target.Value = sourceData.ToString("yyyy-MM-dd HH:mm:ss");

                    // reset
                    numConsecutiveFailures = 0;
                }
                // If we can't read into a DateTime object, just skip it.
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // An occasional miss is ok, but three in a row & we've run outta data.
                    numConsecutiveFailures++;

                    if (numConsecutiveFailures >= 3)
                    {
                        break;
                    }
                }
            }
        }
    }
}
