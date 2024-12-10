using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    /// <summary>
    /// Which part of the survey are we in?
    /// </summary>
    internal enum SurveySection
    {
        [Description("Medical Practice")]
        MedicalPractice,

        [Description("Telehealth")]
        Telehealth,

        [Description("Unknown")]
        Unknown
    }

    internal class SurveyRow
    {
        private readonly string provider;
        private readonly SurveySection section;
        private const string QUOTE = "'";
        private List<string> payload;

        internal SurveyRow(string _provider, SurveySection _section, Range target)
        {
            provider = _provider;
            section = _section;

            payload = new List<string>();
            string cellContents;
            int colOffset = 0;

            // Scan this row.
            while (true)
            {
                try
                {
                    cellContents = target.Offset[0, colOffset].Value.ToString();
                    payload.Add(QUOTE + cellContents + QUOTE);
                    colOffset++;
                }
                catch (RuntimeBinderException)
                {
                    // Then we've run out of data.
                    break;
                }
            }
        }

        /// <summary>
        /// What's the result?
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return "(" + QUOTE + provider + QUOTE + ", " +
                QUOTE + section.ToString() + QUOTE + ", " +
                string.Join(",", payload) + ")";
        }
    }

    internal class SurveyResults
    {
        private Dictionary<string, SurveySection> surveySectionDictionary;

        private const string PREAMBLE = "USE [REL_CLARITY];\r\n\r\n";
        private const string SEGMENT_START = "INSERT INTO #PATIENT_SATISFACTION_LIST (PROVIDER_NAME, SECTION_NAME, QUESTION, SCORE, RANK_NUM, NUM_ANSWERS)\r\nVALUES\r\n";

        internal SurveyResults()
        {
            // Initialize needed dictionary.
            InitializeSurveySectionDictionary();
        }

        private void InitializeSurveySectionDictionary()
        {
            surveySectionDictionary = new Dictionary<string, SurveySection>();

            // Get all the values.
            SurveySection[] surveySections = (SurveySection[])Enum.GetValues(typeof(SurveySection));

            foreach (SurveySection surveySection in surveySections)
            {
                surveySectionDictionary.Add(surveySection.GetDescription(), surveySection);
            }
        }

        internal void Scan(Worksheet worksheet)
        {
            // Initialize scan.
            Range target = (Range)worksheet.Cells[5,1];
            string provider = string.Empty;
            SurveySection section = SurveySection.Unknown;
            Regex nameDetector = new Regex(@"\w+,(\s*\w*\.?)+,\s*(?:DO|MD)");

            // Initialize the output .SQL file.
            Workbook workbook = worksheet.Parent;
            string workbookFilename = workbook.FullName;

            (StreamWriter writer, string outputFilename) = Utilities.OpenOutput(
                inputFilename: workbookFilename,
                filenameAddon: "_list",
                filetype: ".sql"
            );

            writer.Write(PREAMBLE + SEGMENT_START);

            int rowOffset = 0;
            string contentsFirstCell = string.Empty;
            string line_ending = ",\r\n";
            bool firstRow = true;

            while (true)
            {
                rowOffset++;

                try
                {
                    contentsFirstCell = target.Offset[rowOffset, 0].Value.ToString();
                }
                catch (RuntimeBinderException)
                {
                    // Then we've run out of data, so leave the loop.
                    break;
                }

                Match nameMatch = nameDetector.Match(contentsFirstCell);

                // Does this cell look like a provider's name?
                if (nameMatch.Success)
                {
                    // Then log the new provider name & continue to next row.
                    provider = contentsFirstCell;
                    continue;
                }

                // Does this cell look like a section name?
                try
                {
                    // Then log the new section name & continue to next row.
                    section = surveySectionDictionary[contentsFirstCell];
                    continue;
                }
                catch (KeyNotFoundException) { }

                SurveyRow surveyRow = new SurveyRow(provider, section, target.Offset[rowOffset, 0]);

                // Convert SurveyRow object to string.
                string rowContents = surveyRow.ToString();

                // Do we need to add a line ending to the previous row
                // before writing out this new row?
                if (firstRow)
                {
                    writer.Write(rowContents);
                    firstRow = false;
                }
                else
                {
                    writer.Write(line_ending + rowContents);
                }
            }

            // Last row gets a different ending.
            writer.Write(";\r\n");
            writer.Close();

            // Show the resulting SQL file.
            Process.Start(outputFilename);
        }
    }
}
