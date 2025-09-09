using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Class that pings UCSD Blink site to find providers by their email address.
     */
    internal class EmailSearcher
    {
        private Microsoft.Office.Interop.Excel.Application externalFileApplication;
        private const string emailExtractor = @"(?<name>[^@]+)@";
        private string existingNamesFile = string.Empty;
        private Dictionary<string, string> existingNamesList;
        private HtmlWeb web;
        //private string result = null;
        //private WebResponse response = null;
        //private StreamReader reader = null;
        private Range providerEmailRng;
        private bool? useExistingNamesList = null;

        internal EmailSearcher()
        {
            existingNamesList = new Dictionary<string, string>();

            // https://stackoverflow.com/a/48930280/18749636
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        }

        /// <summary>
        /// Asks the user to point to the existing spreadsheet list of names/emails.
        /// </summary>
        private void AskUserForExternalFile()
        {
            if (string.IsNullOrEmpty(existingNamesFile))
            {
                // Have we already asked the user if there's an
                // existing spreadsheet with emails & names?
                if (useExistingNamesList is null)
                {
                    useExistingNamesList = AskUserIfUsingExistingList();
                }

                if (useExistingNamesList.HasValue &&
                    useExistingNamesList.Value)
                {
                    using (OpenFileDialog openFileDialog = new OpenFileDialog())
                    {
                        // Because we don't specify an opening directory,
                        // the dialog will open in the last directory used.
                        openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                        openFileDialog.RestoreDirectory = true;

                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            // Get the path of specified file.
                            existingNamesFile = openFileDialog.FileName;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Asks the user if we're using an existing spreadsheet list of names/emails.
        /// </summary>
        /// <returns>bool</returns>
        private bool AskUserIfUsingExistingList()
        {
            DialogResult dialogResult = MessageBox.Show("Should we look up names in an existing spreadsheet?",
                                                        "External Lookup", MessageBoxButtons.YesNo);
            return dialogResult == DialogResult.Yes;
        }

        /// <summary>
        /// Asks the user if we're using an existing spreadsheet list of names/emails.
        /// </summary>
        /// <param name="namesAndEmails">List<Range></Range></param>
        private void BuildDictionaryFromExternalFile(List<Range> namesAndEmails)
        {
            Range emails = null;
            Range names = null;
            Range firstNames = null;

            switch (namesAndEmails.Count)
            {
                // Maybe it's full name & email.
                case 2:

                    // Which column contains the '@' character?
                    if (Utilities.PresentInColumn(namesAndEmails[0], "@"))
                    {
                        // Then first range is emails.
                        emails = namesAndEmails[0];
                        names = namesAndEmails[1];
                    }
                    else if (Utilities.PresentInColumn(namesAndEmails[1], "@"))
                    {
                        emails = namesAndEmails[1];
                        names = namesAndEmails[0];
                    }
                    else
                    {
                        string message = "Unable to find email column containing '@'.";
                        string title = "Not Found";
                        MessageBoxButtons buttons = MessageBoxButtons.OK;
                        DialogResult result = MessageBox.Show(message, title, buttons, MessageBoxIcon.Warning);

                        if (result == DialogResult.OK)
                        {
                            return;
                        }
                    }

                    break;

                // Maybe it's first name, last name & email.
                case 3:

                    List<int> possibleColumnNumbers = new List<int> { 0, 1, 2 };

                    // Which column contains the '@' character?
                    if (Utilities.PresentInColumn(namesAndEmails[0], "@"))
                    {
                        // Then first range is emails.
                        emails = namesAndEmails[0];
                        possibleColumnNumbers.Remove(0);
                    }
                    else if (Utilities.PresentInColumn(namesAndEmails[1], "@"))
                    {
                        emails = namesAndEmails[1];
                        possibleColumnNumbers.Remove(1);
                    }
                    else if (Utilities.PresentInColumn(namesAndEmails[2], "@"))
                    {
                        emails = namesAndEmails[2];
                        possibleColumnNumbers.Remove(2);
                    }
                    else
                    {
                        string message = "Unable to find email column containing '@'.";
                        string title = "Not Found";
                        MessageBoxButtons buttons = MessageBoxButtons.OK;
                        DialogResult result = MessageBox.Show(message, title, buttons, MessageBoxIcon.Warning);

                        if (result == DialogResult.OK)
                        {
                            return;
                        }
                    }

                    // Figure out firstname/lastname from row 1 headers.
                    foreach (int colNumber in possibleColumnNumbers)
                    {
                        Range thisColumn = namesAndEmails[colNumber];

                        if (Utilities.PresentInHeader(thisColumn, "First"))
                        {
                            firstNames = thisColumn;
                            List<int> remainingColumnNumbers = new List<int>(possibleColumnNumbers);
                            remainingColumnNumbers.Remove(colNumber);
                            names = namesAndEmails[remainingColumnNumbers[0]];
                            break;
                        }
                    }

                    break;

                default:
                    return;
            }

            ScrapeNamesAndEmails(emails, names, firstNames);

            // Close the external spreadsheet.
            externalFileApplication.Quit();
        }

        /// <summary>
        /// Extracts everything before the "@" symbol.
        /// </summary>
        /// <param name="emailRng">Range</param>
        /// <returns>string</returns>
        private string ExtractJustNameFromEmail(Range emailRng)
        {
            string result = string.Empty;

            try
            {
                string email = Convert.ToString(emailRng.Value2);

                if (!string.IsNullOrEmpty(email))
                {
                    Match match = Regex.Match(email, emailExtractor);

                    if (match.Success)
                    {
                        result = Convert.ToString(match.Groups["name"].Value);
                    }
                }
            }
            catch (System.NullReferenceException) { }

            return result;
        }

        /// <summary>
        /// Extracts everything before the "@" symbol.
        /// </summary>
        /// <param name="emailRng">Range</param>
        /// <returns>string</returns>
        private string ExtractJustNameFromEmail(string email)
        {
            string result = string.Empty;

            if (!string.IsNullOrEmpty(email))
            {
                Match match = Regex.Match(email, emailExtractor);

                if (match.Success)
                {
                    result = Convert.ToString(match.Groups["name"].Value);
                }
            }

            return result;
        }

        /// <summary>
        /// Pulls just the provider name from "Faculty/Staff Directory: Doctor, Ima"
        /// </summary>
        /// <param name="titleAndName">string</param>
        /// <returns></returns>
        private string ExtractJustTheNameFromTitle(string titleAndName)
        {
            string result = string.Empty;
            string titleExtractor = @"Staff Directory: (?<name>[\w\s'-]+,[\w\s'-]+)";

            if (!string.IsNullOrEmpty(titleAndName))
            {
                // Substitute html &#039; with apostrophe.
                titleAndName = titleAndName.Trim().Replace("&#039;", "'");
                Match match = Regex.Match(titleAndName, titleExtractor);

                if (match.Success)
                {
                    result = Convert.ToString(match.Groups["name"].Value);
                }
            }

            return result;
        }

        /// <summary>
        /// Finds the best match for an email name in a list of emails/names.
        /// </summary>
        /// <param name="emailName">string</param>
        /// <param name="nameNodes">HtmlNodeCollection</param>
        /// <param name="emailNodes">HtmlNodeCollection</param>
        /// <returns>string</returns>

        private string FindBestMatch(string emailName,
                                     HtmlAgilityPack.HtmlNodeCollection nameNodes,
                                     HtmlAgilityPack.HtmlNodeCollection emailNodes)
        {
            string name = string.Empty;

            if (nameNodes != null && emailNodes != null)
            {
                int numNodes = Math.Min(nameNodes.Count(), emailNodes.Count());

                for (int index = 0; index < numNodes; index++)
                {
                    // Look for an exact match.
                    string thisEmail = ExtractJustNameFromEmail(emailNodes[index].InnerText.Trim());

                    if (thisEmail == emailName)
                    {
                        name = nameNodes[index].InnerText.Trim();
                        break;
                    }
                }
            }

            return name;
        }

        /// <summary>
        /// Finds the column the user has selected. If none, asks them to select from a list.
        /// </summary>
        /// <param name="worksheet">Worksheet</param>
        /// <returns>bool</returns>

        private bool FindSelectedCategory(Worksheet worksheet)
        {
            bool success = false;
            int lastRowInSheet = worksheet.UsedRange.Rows.Count;
            Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.Application;

            // Any column selected?
            providerEmailRng = Utilities.GetSelectedCol(application);

            if (providerEmailRng is null)
            {
                // Then ask user to select one column.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames, MultiSelect: false))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        string selectedColumnName = form.selectedColumns[0];
                        providerEmailRng = Utilities.TopOfNamedColumn(worksheet, selectedColumnName);
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
        /// Finds the user-selected columns in an external spreadsheet.
        /// </summary>
        /// <param name="filename">string</param>
        /// <returns>List<Range></Range></returns>
        private List<Range> FindSelectedColumns(string filename)
        {
            List<Range> selectedColumnsRng = new List<Range>();

            Workbook workbook = Utilities.OpenExternalFile(filename);

            if (workbook != null)
            {
                externalFileApplication = workbook.Application;

                // Tell me when you're ready.
                DialogResult dialogResult = MessageBox.Show("Click 'OK' once you've selected name & email columns.",
                                                            "Please select name & email columns.",
                                                            MessageBoxButtons.OKCancel,
                                                            MessageBoxIcon.Question,
                                                            MessageBoxDefaultButton.Button1,
                                                            MessageBoxOptions.DefaultDesktopOnly);

                if (dialogResult == DialogResult.OK)
                {
                    // Any column selected?
                    int lastRow = workbook.ActiveSheet.UsedRange.Rows.Count;
                    selectedColumnsRng = Utilities.GetSelectedCols(externalFileApplication);
                }
                else
                {
                    // If user pressed "Cancel", then don't ask again.
                    useExistingNamesList = false;
                }
            }

            return selectedColumnsRng;
        }

        /// <summary>
        /// Looks up the user's name in an external list of emails/names.
        /// </summary>
        /// <param name="emailName">string</param>
        /// <returns>string</returns>
        private string LookupNameOutside(string emailName)
        {
            string result = string.Empty;

            if (existingNamesList.Count > 0)
            {
                try
                {
                    result = existingNamesList[emailName];
                }
                catch (KeyNotFoundException) { }
            }
            else
            {
                AskUserForExternalFile();

                if (!string.IsNullOrEmpty(existingNamesFile))
                {
                    List<Range> externalNameRange = FindSelectedColumns(existingNamesFile);
                    BuildDictionaryFromExternalFile(externalNameRange);
                    result = LookupNameOutside(emailName);
                }
            }

            return result;
        }

        /// <summary>
        /// Pings the UCSD Blink service to look up the user's name from their email.
        /// </summary>
        /// <param name="emailName">string</param>
        /// <returns>string</returns>
        private string PingUcsdBlink(string emailName)
        {
            string result = string.Empty;
            string url = "https://itsweb.ucsd.edu/directory/search?t=directory&entry=";
            string urlPopulated = url + emailName;
            HtmlAgilityPack.HtmlDocument doc = web.Load(urlPopulated);

            // Is there just one match?
            try
            {
                result = doc.GetElementbyId("empNameTitle").InnerText;
            }
            catch (System.NullReferenceException)
            {
                // Or multiple matches?
                try
                {
                    HtmlAgilityPack.HtmlNodeCollection nameNodes = doc.DocumentNode.SelectNodes("//table/tbody/tr/td[1]/a");
                    HtmlAgilityPack.HtmlNodeCollection emailNodes = doc.DocumentNode.SelectNodes("//table/tbody/tr/td[5]/a");
                    return FindBestMatch(emailName, nameNodes, emailNodes);
                }
                catch (System.NullReferenceException) { }
            }

            return ExtractJustTheNameFromTitle(result);
        }

        /// <summary>
        /// Reads/parses the emails addresses and names in an external file
        /// to build a Dictionary.
        /// </summary>
        /// <param name="emailsColumn">range</param>
        /// <param name="namesColumn">range</param>
        private void ScrapeNamesAndEmails(Range emailsColumn, Range namesColumn, Range firstNamesColumn = null)
        {
            Worksheet externalWorksheet = emailsColumn.Parent;
            int lastRowInExternalSheet = externalWorksheet.UsedRange.Rows.Count;

            Range topOfEmails = (Range)externalWorksheet.Cells[1, emailsColumn.Column];
            Range topOfNames = (Range)externalWorksheet.Cells[1, namesColumn.Column];

            bool usingSeparateNames = false;
            Range topOfFirstNames = null;

            if (firstNamesColumn != null)
            {
                usingSeparateNames = true;
                topOfFirstNames = (Range)externalWorksheet.Cells[1, firstNamesColumn.Column];
            }

            for (int rowOffset = 1; rowOffset <= lastRowInExternalSheet; rowOffset++)
            {
                string emailAddress = string.Empty;
                string emailName = string.Empty;
                string name = string.Empty;
                string firstName = string.Empty;

                try
                {
                    emailAddress = topOfEmails.Offset[rowOffset, 0].Value.ToString();
                    emailName = ExtractJustNameFromEmail(emailAddress);
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    return;
                }

                try
                {
                    name = topOfNames.Offset[rowOffset, 0].Value.ToString();
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    return;
                }

                if (usingSeparateNames)
                {
                    try
                    {
                        firstName = topOfFirstNames.Offset[rowOffset, 0].Value.ToString();
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        return;
                    }
                }

                if (!string.IsNullOrEmpty(firstName))
                {
                    name = name + ", " + firstName;
                }

                if (!string.IsNullOrEmpty(emailName) && !string.IsNullOrEmpty(name))
                {
                    existingNamesList.Add(emailName, name);
                }
            }
        }

        /// <summary>Main routine</summary>
        /// <param name="worksheet">Worksheet</param>
        internal void Search(Worksheet worksheet)
        {
            int lastRowInSheet = worksheet.UsedRange.Rows.Count;

            if (FindSelectedCategory(worksheet))
            {
                // Create object here--once.
                web = new HtmlWeb();

                // Create column for provider name.
                Range providerNameRng = Utilities.InsertNewColumn(providerEmailRng, "Provider Name");

                // Run down the email column, pinging UCSD Blink & parsing out their name.
                for (int rowOffset = 1; rowOffset < lastRowInSheet; rowOffset++)
                {
                    string emailName = ExtractJustNameFromEmail(providerEmailRng.Offset[rowOffset]);

                    // Were we able to parse "improvider" from "improvider@health.ucsd.edu"?
                    if (!string.IsNullOrEmpty(emailName))
                    {
                        // Ask the Blink server.
                        string providerName = PingUcsdBlink(emailName);

                        // If that didn't work, look it up in an existing spreadsheet.
                        if (string.IsNullOrEmpty(providerName))
                        {
                            providerName = LookupNameOutside(emailName);
                        }

                        providerNameRng.Offset[rowOffset].Value = providerName;
                    }
                }
            }
        }
    }
}