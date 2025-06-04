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
        private int lastRow;
        private Range selectedColumnRng;
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

        private bool FindSelectedColumn(Worksheet worksheet)
        {
            bool success = false;
            lastRow = worksheet.UsedRange.Rows.Count;

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
            int lastRow = worksheet.UsedRange.Rows.Count;

            if (FindSelectedColumn(worksheet))
            {
                string selectedColumnName = selectedColumnRng.Value.ToString();
                string newColumnName = selectedColumnName + " Text";

                // Make room for new column.
                Range ditheredColumn = Utilities.InsertNewColumn(range: selectedColumnRng,
                                                                 newColumnName: newColumnName,
                                                                 side: InsertSide.Right);
                ditheredColumn.NumberFormat = "@";

                DateTime sourceData;
                Range target;

                for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
                {
                    target = (Range)worksheet.Cells[rowNumber, ditheredColumn.Column];

                    try
                    {
                        sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value;
                        target.Value = sourceData.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    // If we can't read into a DateTime object, just skip it.
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                }
            }
        }
    }
}
