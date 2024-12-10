using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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

    /// <summary>
    /// Questions re: medical practice
    /// </summary>
    internal enum MedicalPracticeQuestions
    {
        [Description("Key Metric NPS: Provider would recommend")]
        WouldRecommend,

        [Description("Able to get appt")]
        AbleToGetAppt,

        [Description("Got enough info re: treatment")]
        GotEnoughInfo,

        [Description("Informed of delays")]
        InformedOfDelays,

        [Description("Knew what to do if questions")]
        KnewWhatToDoIfQuestions,

        [Description("Office staff courteous/helpful")]
        StaffHelpful,

        [Description("Trust provider w/ care")]
        TrustProvider,

        [Description("Treated respectfully w/o bias")]
        TreatedRespectfully,

        [Description("Provider timely to see you")]
        SeenTimely,

        [Description("Unknown")]
        Unknown
    }

    /// <summary>
    /// Questions re: telehealth
    /// </summary>
    internal enum TelehealthQuestions
    {
        [Description("Key Metric NPS: Provider would recommend")]
        WouldRecommend,

        [Description("Trust provider w/ care")]
        TrustProvider,

        [Description("Treated respectfully w/o bias")]
        TreatedRespectfully,

        [Description("Able to get appt")]
        AbleToGetAppt,

        [Description("Provider timely to see you")]
        SeenTimely,

        [Description("Informed of delays")]
        InformedOfDelays,

        [Description("Gave enough info")]
        GaveEnoughInfo,

        [Description("Knew what to do if questions")]
        KnewWhatToDoIfQuestions,

        [Description("Method of connecting was easy")]
        EasyToConnect,

        [Description("Unknown")]
        Unknown
    }

    /// <summary>
    /// Wraps up a medical practice question with its scores.
    /// </summary>
    internal class MedicalQuestionResults
    {
        private readonly MedicalPracticeQuestions question;
        private readonly double boxScore;
        private readonly int percentileRank;
        private readonly int responseSize;

        internal MedicalQuestionResults(MedicalPracticeQuestions _question, double _boxScore, int _rank, int _size)
        {
            question = _question;
            boxScore = _boxScore;
            percentileRank = _rank;
            responseSize = _size;
        }

        internal string QuestionName()
        {
            return question.GetDescription();
        }

        internal int Rank() { return percentileRank; }
        internal int ResponseSize() { return responseSize; }
        internal double Score() { return boxScore; }
    }

    internal class SurveyRow
    {
        private readonly string provider;
        private readonly SurveySection section;
        private readonly string question;
        private readonly double score;
        private readonly int rank;
        private readonly int size;
        private const string QUOTE = "'";

        internal SurveyRow(string _provider, MedicalQuestionResults results)
        {
            provider = _provider;
            section = SurveySection.MedicalPractice;
            question = results.QuestionName();
            score = results.Score();
            rank = results.Rank();  
            size = results.ResponseSize();
        }

        internal SurveyRow(string _provider, TelehealthQuestionResults results)
        {
            provider = _provider;
            section = SurveySection.Telehealth;
            question = results.QuestionName();
            score = results.Score();
            rank = results.Rank();
            size = results.ResponseSize();
        }

        /// <summary>
        /// What's the result?
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return "(" + QUOTE + provider + QUOTE + ", " +
                QUOTE + section.ToString() + QUOTE + ", " +
                QUOTE + question + QUOTE + ", " +
                QUOTE + score.ToString() + QUOTE + ", " +
                QUOTE + rank.ToString() + QUOTE + ", " +
                QUOTE + size.ToString() + QUOTE + ")";
        }
    }

    /// <summary>
    /// Wraps up a telehealth question with its scores.
    /// </summary>
    internal class TelehealthQuestionResults
    {
        private readonly TelehealthQuestions question;
        private readonly double boxScore;
        private readonly int percentileRank;
        private readonly int responseSize;

        internal TelehealthQuestionResults(TelehealthQuestions _question, double _boxScore, int _rank, int _size)
        {
            question = _question;
            boxScore = _boxScore;
            percentileRank = _rank;
            responseSize = _size;
        }

        internal string QuestionName()
        {
            return question.GetDescription();
        }
        internal int Rank() { return percentileRank; }
        internal int ResponseSize() { return responseSize; }
        internal double Score() { return boxScore; }
    }

    internal class SurveyResults
    {
        private List<MedicalQuestionResults> medicalResults;
        private List<TelehealthQuestions> telehealthResults;
        private Dictionary<string, MedicalPracticeQuestions> medicalQuestionDictionary;
        private Dictionary<string, SurveySection> surveySectionDictionary;
        private Dictionary<string, TelehealthQuestions> telehealthQuestionDictionary;

        private const string PREAMBLE = "USE [REL_CLARITY];\r\n\r\n";
        private const string SEGMENT_START = "INSERT INTO #PATIENT_SATISFACTION_LIST (PROVIDER_NAME, SECTION_NAME, QUESTION, SCORE, RANK_NUM, NUM_ANSWERS)\r\nVALUES\r\n";

        internal SurveyResults()
        {
            // Initialize needed dictionaries.
            InitializeSurveySectionDictionary();
            InitializeMedicalQuestionDictionary();
            InitializeTelehealthQuestionDictionary();
        }

        /// <summary>
        /// Build a list of the cell contents across this row.
        /// </summary>
        /// <param name="target">Range of first cell in row to be parsed</param>
        /// <returns>MedicalQuestionResults</returns>
        private MedicalQuestionResults ExtractMedicalPracticeRow(Range target)
        {
            string cellContents;
            MedicalQuestionResults results = null;
            MedicalPracticeQuestions question = MedicalPracticeQuestions.Unknown;
            double? score = null;
            int? rank = null;
            int? size = null;

            // We expect to see question, score, percentile and size.
            try
            {
                cellContents = Convert.ToString(target.Value2);

                try
                {
                    question = medicalQuestionDictionary[cellContents];
                }
                catch (KeyNotFoundException)
                {
                    return results;
                }

                cellContents = Convert.ToString(target.Offset[0, 1].Value2);

                if (double.TryParse(cellContents, out double _score))
                {
                    score = _score;
                }

                cellContents = Convert.ToString(target.Offset[0, 2].Value2);

                if (int.TryParse(cellContents, out int _rank))
                {
                    rank = _rank;
                }

                cellContents = Convert.ToString(target.Offset[0, 3].Value2);

                if (int.TryParse(cellContents, out int _size))
                {
                    size = _size;
                }

                if (score.HasValue && rank.HasValue && size.HasValue)
                {
                    results = new MedicalQuestionResults(question, score.Value, rank.Value, size.Value);
                }
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }

            return results;
        }

        /// <summary>
        /// Build a list of the cell contents across this row.
        /// </summary>
        /// <param name="target">Range of first cell in row to be parsed</param>
        /// <returns>TelehealthQuestionResults</returns>
        private TelehealthQuestionResults ExtractTelehealthRow(Range target)
        {
            string cellContents;
            TelehealthQuestionResults results = null;
            TelehealthQuestions question = TelehealthQuestions.Unknown;
            double? score = null;
            int? rank = null;
            int? size = null;

            // We expect to see question, score, percentile and size.
            try
            {
                cellContents = Convert.ToString(target.Value2);

                try
                {
                    question = telehealthQuestionDictionary[cellContents];
                }
                catch (KeyNotFoundException)
                {
                    return results;
                }

                cellContents = Convert.ToString(target.Offset[0, 1].Value2);

                if (double.TryParse(cellContents, out double _score))
                {
                    score = _score;
                }

                cellContents = Convert.ToString(target.Offset[0, 2].Value2);

                if (int.TryParse(cellContents, out int _rank))
                {
                    rank = _rank;
                }

                cellContents = Convert.ToString(target.Offset[0, 3].Value2);

                if (int.TryParse(cellContents, out int _size))
                {
                    size = _size;
                }

                if (score.HasValue && rank.HasValue && size.HasValue)
                {
                    results = new TelehealthQuestionResults(question, score.Value, rank.Value, size.Value);
                }
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }

            return results;
        }

        private void InitializeMedicalQuestionDictionary()
        {
            medicalQuestionDictionary = new Dictionary<string, MedicalPracticeQuestions>();

            // Get all the values.
            MedicalPracticeQuestions[] questions = (MedicalPracticeQuestions[])Enum.GetValues(typeof(MedicalPracticeQuestions));

            foreach (MedicalPracticeQuestions question in questions)
            {
                medicalQuestionDictionary.Add(question.GetDescription(), question);
            }
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

        private void InitializeTelehealthQuestionDictionary()
        {
            telehealthQuestionDictionary = new Dictionary<string, TelehealthQuestions>();

            // Get all the values.
            TelehealthQuestions[] questions = (TelehealthQuestions[])Enum.GetValues(typeof(TelehealthQuestions));

            foreach (TelehealthQuestions question in questions)
            {
                telehealthQuestionDictionary.Add(question.GetDescription(), question);
            }
        }

        internal void Scan(Worksheet worksheet)
        {
            // Initialize scan.
            Range target = (Range)worksheet.Cells[6,1];
            string provider = string.Empty;
            SurveySection section = SurveySection.Unknown;
            bool haveProvider = false;
            bool haveSection = false;
            Regex nameDetector = new Regex(@"\w+,\s*\w+\s*\w*\.?,\s*(DO|MD)");

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
                try
                {
                    contentsFirstCell = target.Offset[rowOffset, 0].Value.ToString();
                }
                catch (RuntimeBinderException)
                {
                    // Then we've run out of data.
                    break;
                }

                // Keep looking for provider until we find one.
                if (!haveProvider) 
                {
                    Match nameMatch = nameDetector.Match(contentsFirstCell);

                    if (nameMatch.Success)
                    {
                        provider = contentsFirstCell;
                        haveProvider = true;
                    }

                    rowOffset++;
                    continue;
                }

                // Do we have a section name yet?
                if (!haveSection)
                {
                    try
                    {
                        section = surveySectionDictionary[contentsFirstCell];
                        haveSection = true;
                    }
                    catch(KeyNotFoundException) { }

                    rowOffset++;
                    continue;
                }

                SurveyRow surveyRow;

                if (section == SurveySection.MedicalPractice)
                {
                    MedicalQuestionResults results = ExtractMedicalPracticeRow(target.Offset[rowOffset, 0]);

                    if (results == null)
                    {
                        // Maybe we've moved on to the next section.
                        haveSection = false;
                        continue;
                    }

                    // Create a SurveyRow object.
                    surveyRow = new SurveyRow(provider, results);
                }
                else
                {
                    TelehealthQuestionResults results = ExtractTelehealthRow(target.Offset[rowOffset, 0]);

                    if (results == null)
                    {
                        // Maybe we've moved on to the next provider & section.
                        haveProvider = false;
                        haveSection = false;
                        continue;
                    }

                    // Create a SurveyRow object.
                    surveyRow = new SurveyRow(provider, results);
                }

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

                rowOffset++;
            }

            // Last row gets a different ending.
            writer.Write(";\r\n");
            writer.Close();

            // Show the resulting SQL file.
            Process.Start(outputFilename);
        }
    }
}
