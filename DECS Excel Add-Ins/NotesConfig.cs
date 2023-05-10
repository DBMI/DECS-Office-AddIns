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
        // The Regular ExpressionS to search for...
        public Pattern[] Patterns { get; set; }

        // ...and the new column to be created.
        public string new_column { get; set; }
    }
    // Defines the way the current workbook & sheet should be parsed.
    public class NotesConfig
    {
        public string SourceColumn { get; set; }
        public CleaningRule[] CleaningRules { get; set; }

        public ExtractRule[] ExtractRules { get; set; }

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
    public class Pattern
    {
        public string pattern { get; set;}
    }
}
