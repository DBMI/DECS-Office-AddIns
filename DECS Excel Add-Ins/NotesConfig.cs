using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using log4net;

namespace DECS_Excel_Add_Ins
{
    // Defines a single replacement rule.
    public class CleaningRule
    {
        public string displayName { get; set; }

        public bool enabled { get; set; }

        // The Regular Expression to search for...
        public string pattern { get; set; }

        // ...and what to replace it with.
        public string replace { get; set; }

        public CleaningRule()
        {
            enabled = false;
        }
    }

    public class DateConversionRule
    {
        public bool enabled { get; set; }

        public string desiredDateFormat { get; set; }

        public DateConversionRule()
        {
            enabled = false;
        }
    }

    //  Defines a single extraction rule.
    public class ExtractRule
    {
        public string displayName { get; set; }

        public bool enabled { get; set; }

        // The Regular Expression to search for...
        public string pattern { get; set; }

        // ...and the new column to be created.
        public string newColumn { get; set; }

        public ExtractRule()
        {
            enabled = false;
        }
    }

    // Defines the way the current workbook & sheet should be parsed.
    public class NotesConfig
    {
        public List<CleaningRule> CleaningRules { get; set; }

        public DateConversionRule DateConversionRule { get; set; }

        public List<ExtractRule> ExtractRules { get; set; }

        public string SourceColumnName { get; set; }

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        // Constructor
        internal NotesConfig()
        {
            log.Debug("Instantiating a NotesConfig object.");
            SourceColumnName = string.Empty;
            CleaningRules = new List<CleaningRule>();
            DateConversionRule = new DateConversionRule();
            ExtractRules = new List<ExtractRule>();
        }

        internal void AddCleaningRule()
        {
            CleaningRules.Add(new CleaningRule());
        }

        internal void AddExtractRule()
        {
            ExtractRules.Add(new ExtractRule());
        }

        internal void ChangeCleaningRuleDisplayName(int index, string displayName)
        {
            if (CleaningRules.Count - 1 >= index)
            {
                CleaningRules[index].displayName = displayName;
            }
        }

        internal void ChangeCleaningRulePattern(int index, string pattern)
        {
            if (CleaningRules.Count - 1 >= index)
            {
                CleaningRules[index].pattern = pattern;
            }
        }

        internal void ChangeCleaningRuleReplace(int index, string replace)
        {
            if (CleaningRules.Count - 1 >= index)
            {
                CleaningRules[index].replace = replace;
            }
        }

        internal void ChangeExtractRuleDisplayName(int index, string displayName)
        {
            if (ExtractRules.Count - 1 >= index)
            {
                ExtractRules[index].displayName = displayName;
            }
        }

        internal void ChangeExtractRulePattern(int index, string pattern)
        {
            if (ExtractRules.Count - 1 >= index)
            {
                ExtractRules[index].pattern = pattern;
            }
        }

        internal void ChangeExtractRulenewColumn(int index, string newColumn)
        {
            if (ExtractRules.Count - 1 >= index)
            {
                ExtractRules[index].newColumn = newColumn;
            }
        }

        internal static string ChooseConfigFile()
        {
            string filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // Because we don't specify an opening directory,
                // the dialog will open in the last directory used.
                openFileDialog.Filter = "xml files (*.xml)|*.xml";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of specified file.
                    filePath = openFileDialog.FileName;
                }
            }

            return filePath;
        }

        internal void DeleteCleaningRule(int index)
        {
            if (index >= 0 && index < CleaningRules.Count)
            {
                CleaningRules.RemoveAt(index);
            }
        }

        internal void DeleteExtractRule(int index)
        {
            if (index >= 0 && index < ExtractRules.Count)
            {
                ExtractRules.RemoveAt(index);
            }
        }

        internal void DisableCleaningRule(int index)
        {
            if (index >= 0 && index < CleaningRules.Count)
            {
                CleaningRules[index].enabled = false;
            }
        }

        internal void DisableExtractRule(int index)
        {
            if (index >= 0 && index < ExtractRules.Count)
            {
                ExtractRules[index].enabled = false;
            }
        }

        internal void EnableCleaningRule(int index)
        {
            if (index >= 0 && index < CleaningRules.Count)
            {
                CleaningRules[index].enabled = true;
            }
        }

        internal void EnableExtractRule(int index)
        {
            if (index >= 0 && index < ExtractRules.Count)
            {
                ExtractRules[index].enabled = true;
            }
        }

        internal bool HasCleaningRules()
        {
            List<CleaningRule> validRules = ValidCleaningRules();
            return validRules.Count > 0;
        }

        internal bool HasDateConversionRule()
        {
            return DateConversionRule.enabled
                && !string.IsNullOrEmpty(DateConversionRule.desiredDateFormat);
        }

        internal bool HasExtractRules()
        {
            List<ExtractRule> validRules = ValidExtractRules();
            return validRules.Count > 0;
        }

        internal static NotesConfig ReadConfigFile(string filePath)
        {
            // Declare this outside the 'using' block so we can access it later
            NotesConfig config = null;

            if (string.IsNullOrEmpty(filePath))
            {
                return config;
            }

            using (var reader = new StreamReader(filePath))
            {
                config = (NotesConfig)new XmlSerializer(typeof(NotesConfig)).Deserialize(reader);
            }

            return config;
        }

        internal int NumValidCleaningRules()
        {
            List<CleaningRule> validRules = ValidCleaningRules();
            return validRules.Count;
        }

        internal int NumValidExternalRules()
        {
            List<ExtractRule> validRules = ValidExtractRules();
            return validRules.Count;
        }

        internal List<CleaningRule> ValidCleaningRules()
        {
            List<CleaningRule> validRules = CleaningRules
                .Where(r => r.pattern != null && r.replace != null && r.enabled)
                .ToList();
            return validRules;
        }

        internal List<ExtractRule> ValidExtractRules()
        {
            List<ExtractRule> validRules = ExtractRules
                .Where(r => r.pattern != null && r.newColumn != null && r.enabled)
                .ToList();
            return validRules;
        }

        internal List<RuleValidationError> ValidateRules()
        {
            List<RuleValidationError> errorReports = new List<RuleValidationError>();

            for (int index = 0; index < CleaningRules.Count; index++)
            {
                CleaningRule rule = CleaningRules[index];
                RuleValidationResult result = Utilities.IsRegexValid(rule.pattern);

                if (!result.Valid())
                {
                    RuleValidationError ruleValidationError = new RuleValidationError(
                        _ruleType: RuleType.Cleaning,
                        _index: index,
                        _ruleComponent: RuleComponent.Pattern,
                        _message: result.ToString()
                    );
                    errorReports.Add(ruleValidationError);
                }

                result = Utilities.IsRegexValid(rule.replace);

                if (!result.Valid())
                {
                    RuleValidationError ruleValidationError = new RuleValidationError(
                        _ruleType: RuleType.Cleaning,
                        _index: index,
                        _ruleComponent: RuleComponent.Replace,
                        _message: result.ToString()
                    );
                    errorReports.Add(ruleValidationError);
                }
            }

            for (int index = 0; index < ExtractRules.Count; index++)
            {
                ExtractRule rule = ExtractRules[index];
                RuleValidationResult result = Utilities.IsRegexValid(rule.pattern);

                if (!result.Valid())
                {
                    RuleValidationError ruleValidationError = new RuleValidationError(
                        _ruleType: RuleType.Extract,
                        _index: index,
                        _ruleComponent: RuleComponent.Pattern,
                        _message: result.ToString()
                    );
                    errorReports.Add(ruleValidationError);
                }

                if (string.IsNullOrEmpty(rule.newColumn))
                {
                    RuleValidationError ruleValidationError = new RuleValidationError(
                        _ruleType: RuleType.Extract,
                        _index: index,
                        _ruleComponent: RuleComponent.NewColumn,
                        _message: "newColumn is empty."
                    );
                    errorReports.Add(ruleValidationError);
                }
            }

            return errorReports;
        }
    }
}
