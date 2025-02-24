using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using HtmlAgilityPack;
using System.Security.Policy;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Class that pings UCSD Blink site to find providers by their email address.
     */
    internal class EmailSearcher
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private const string emailExtractor = @"(?<name>[^@]+)@";
        private const string titleExtractor = @"Staff Directory: (?<name>[\w\s'-]+,[\w\s'-]+)";
        private HtmlWeb web;
        private int lastRowInSheet;
        private Range providerNameRng;
        private string result = null;
        private WebResponse response = null;
        private StreamReader reader = null;
        private Range providerEmailRng;
        private const string url = "https://itsweb.ucsd.edu/directory/search?t=directory&entry=";

        internal EmailSearcher()
        {
            application = Globals.ThisAddIn.Application;

            // https://stackoverflow.com/a/48930280/18749636
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
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

        private bool FindSelectedCategory(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            providerEmailRng = Utilities.GetSelectedCol(application, lastRowInSheet);

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

        private string PingUcsdBlink(string emailName)
        {
            string result = string.Empty;
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

        internal void Search(Worksheet worksheet)
        {
            lastRowInSheet = worksheet.UsedRange.Rows.Count;

            if (FindSelectedCategory(worksheet))
            {
                web = new HtmlWeb();

                // Create column for provider name.
                providerNameRng = Utilities.InsertNewColumn(providerEmailRng, "Provider Name");

                // Run down the email column, pinging UCSD Blink & parsing out their name.
                for (int rowOffset = 1; rowOffset < lastRowInSheet; rowOffset++)
                {
                    string emailName = ExtractJustNameFromEmail(providerEmailRng.Offset[rowOffset]);

                    if (!string.IsNullOrEmpty(emailName))
                    {
                        string providerName = PingUcsdBlink(emailName);
                        providerNameRng.Offset[rowOffset].Value = providerName;
                    }
                }
            }
        }
    }
}