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

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Class that pings UCSD Blink site to find providers by their email address.
     */
    internal class EmailSearcher
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private const string emailExtractor = @"(?<name>[^@]+)@";
        private const string titleExtractor = @"Staff Directory: (?<name>[\w,\s\.'-]+)";
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

        /// <summary>
        /// Extracts everything before the "@" symbol.
        /// </summary>
        /// <param name="emailRng">Range</param>
        /// <returns>string</returns>
        private string ExtractJustNameFromEmail(Range emailRng)
        {
            string result = string.Empty;
            string email = emailRng.Value2.ToString();

            if (!string.IsNullOrEmpty(email))
            {
                Match match = Regex.Match(email, emailExtractor);

                if (match.Success)
                {
                    result = match.Groups["name"].Value.ToString();
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
                Match match = Regex.Match(titleAndName.Trim(), titleExtractor);

                if (match.Success)
                {
                    result = match.Groups["name"].Value.ToString();
                }
            }

            return result;
        }

        private string PingUcsdBlink(string emailName)
        {
            string result = string.Empty;
            string urlPopulated = url+emailName;
            HtmlAgilityPack.HtmlDocument doc = web.Load(urlPopulated);

            try
            {
                result = doc.GetElementbyId("empNameTitle").InnerText;
            }
            catch (System.NullReferenceException) { }

            return ExtractJustTheNameFromTitle(result);
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

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        string selectedColumnName = form.selectedCategory;
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
    }
}