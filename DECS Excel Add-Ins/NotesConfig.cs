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
        public bool enabled { get; set; }

        // The Regular Expression to search for...
        public string pattern { get; set; }

        // ...and what to replace it with.
        public string replace { get; set; }

        public CleaningRule() 
        {
            enabled = true;
        }
    }
    //  Defines a single extraction rule.
    public class ExtractRule
    {
        public bool enabled { get; set; }

        // The Regular Expression to search for...
        public string pattern { get; set; }

        // ...and the new column to be created.
        public string newColumn { get; set; }

        public ExtractRule() 
        {
            enabled = true;
        }
    }
    // Defines the way the current workbook & sheet should be parsed.
    public class NotesConfig
    {
        public List<CleaningRule> CleaningRules { get; set; }

        public List<ExtractRule> ExtractRules { get; set; }

        public string SourceColumn { get; set; }

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        // Constructor
        internal NotesConfig()
        {
            log.Debug("Instantiating a NotesConfig object.");
            SourceColumn = string.Empty;
            CleaningRules = new List<CleaningRule>();
            ExtractRules = new List<ExtractRule>();
        }
        internal void AddCleaningRule()
        {
            CleaningRules.Add(new CleaningRule());
        }
        internal void AddCleaningRule(CleaningRule cleaningRule)
        {
            CleaningRules.Add(cleaningRule);
        }
        internal void AddExtractRule()
        {
            ExtractRules.Add(new ExtractRule());
        }
        internal void AddExtractRule(ExtractRule extractRule)
        {
            ExtractRules.Add(extractRule);
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
        internal void DisableRule(int index)
        {
            if (index >= 0 && index < CleaningRules.Count)
            {
                CleaningRules[index].enabled = false;
            }

            if (index >= 0 && index < ExtractRules.Count)
            {
                ExtractRules[index].enabled = false;
            }
        }
        internal void EnableRule(int index)
        {
            if (index >= 0 && index < CleaningRules.Count)
            {
                CleaningRules[index].enabled = true;
            }

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
        internal bool HasExtractRules()
        {
            List<ExtractRule> validRules = ValidExtractRules();
            return validRules.Count > 0;
        }
        internal static NotesConfig ReadConfigFile(string filePath)
        {
            // Declare this outside the 'using' block so we can access it later
            NotesConfig config = null;

            if (string.IsNullOrEmpty(filePath)) { return config; }

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
            List<CleaningRule> validRules = CleaningRules.Where(r => r.pattern != null &&
                                                                     r.replace != null &&
                                                                     r.enabled).ToList();
            return validRules;
        }
        internal List<ExtractRule> ValidExtractRules()
        {
            List<ExtractRule> validRules = ExtractRules.Where(r => r.pattern != null &&
                                                                   r.newColumn != null &&
                                                                   r.enabled).ToList();
            return validRules;
        }
        internal List<RuleValidationError> ValidateRules()
        {
            List<RuleValidationError> errorReports = new List<RuleValidationError>();

            for (int index = 0; index < CleaningRules.Count; index++)
            {
                CleaningRule rule = CleaningRules[index];

                if (!Utilities.IsRegexValid(rule.pattern))
                {
                    RuleValidationError ruleValidationError = new RuleValidationError(RuleType.Cleaning, index, RuleComponent.Pattern);
                    errorReports.Add(ruleValidationError);
                }

                if (!Utilities.IsRegexValid(rule.replace))
                {
                    RuleValidationError ruleValidationError = new RuleValidationError(RuleType.Cleaning, index, RuleComponent.Replace);
                    errorReports.Add(ruleValidationError);
                }
            }

            for (int index = 0; index < ExtractRules.Count; index++)
            {
                ExtractRule rule = ExtractRules[index];

                if (!Utilities.IsRegexValid(rule.pattern))
                {
                    RuleValidationError ruleValidationError = new RuleValidationError(RuleType.Extract, index, RuleComponent.Pattern);
                    errorReports.Add(ruleValidationError);
                }

                if (string.IsNullOrEmpty(rule.newColumn))
                {
                    RuleValidationError ruleValidationError = new RuleValidationError(RuleType.Extract, index, RuleComponent.NewColumn);
                    errorReports.Add(ruleValidationError);
                }
            }

            return errorReports;
        }
    }
}