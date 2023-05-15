using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace DECS_Excel_Add_Ins
{
    // Defines a single replacement rule.
    public class CleaningRule
    {
        // The Regular Expression to search for...
        public string pattern { get; set; }

        // ...and what to replace it with.
        public string replace { get; set; }
    }
    //  Defines a single extraction rule.
    public class ExtractRule
    {
        // The Regular Expression to search for...
        public string pattern { get; set; }

        // ...and the new column to be created.
        public string newColumn { get; set; }
    }
    // Defines the way the current workbook & sheet should be parsed.
    public class NotesConfig
    {
        public List<CleaningRule> CleaningRules { get; set; }

        public List<ExtractRule> ExtractRules { get; set; }

        public string SourceColumn { get; set; }

        // Constructor
       internal NotesConfig()
        {
            SourceColumn = string.Empty;
            CleaningRules = new List<CleaningRule>();
            ExtractRules = new List<ExtractRule>();
        }
        public void AddCleaningRule()
        {
            CleaningRules.Add(new CleaningRule());
        }
        public void AddCleaningRule(CleaningRule cleaningRule)
        {
            CleaningRules.Add(cleaningRule);
        }
        public void AddExtractRule()
        {
            ExtractRules.Add(new ExtractRule());
        }
        public void AddExtractRule(ExtractRule extractRule)
        {
            ExtractRules.Add(extractRule);
        }
        public void ChangeCleaningRulePattern(int index, string pattern)
        {
            if (CleaningRules.Count - 1 >= index)
            {
                CleaningRules[index].pattern = pattern;
            }
        }
        public void ChangeCleaningRuleReplace(int index, string replace)
        {
            if (CleaningRules.Count - 1 >= index)
            {
                CleaningRules[index].replace = replace;
            }
        }
        public void ChangeExtractRulePattern(int index, string pattern)
        {
            if (ExtractRules.Count - 1 >= index)
            {
                ExtractRules[index].pattern = pattern;
            }
        }
        public void ChangeExtractRulenewColumn(int index, string newColumn)
        {
            if (ExtractRules.Count - 1 >= index)
            {
                ExtractRules[index].newColumn = newColumn;
            }
        }
        public static string ChooseConfigFile()
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
        public void DeleteCleaningRule(int index)
        {
            if (index > 0 && index < CleaningRules.Count)
            {
                CleaningRules.RemoveAt(index);
            }
        }
        public void DeleteExtractRule(int index)
        {
            if (index > 0 && index < ExtractRules.Count)
            {
                ExtractRules.RemoveAt(index);
            }
        }
        public static NotesConfig ReadConfigFile(string filePath)
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
    }
}
