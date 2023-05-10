using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace DECS_Excel_Add_Ins
{
    internal class NotesParser
    {
        private NotesConfig config;
        private int lastRow;
        private Range sourceColumn;
        private Worksheet worksheet;

        public NotesParser(Worksheet worksheet)
        {
            this.worksheet = worksheet;

            string configFilename = NotesConfig.ChooseConfigFile();
            this.config = NotesConfig.ReadConfigFile(configFilename);

            // If no config already defined, we probably should offer to start defining a config file.
            if (this.config == null) { return; }

            // Find the top of the source column.
            Range rng = Utilities.TopOfNamedColumn(sheet: worksheet, columnName: this.config.SourceColumn);

            if (rng == null)
            {
                Utilities.WarnColumnNotFound(this.config.SourceColumn);
                return; 
            }

            this.sourceColumn = rng;

            // Identify last row.
            this.lastRow = Utilities.FindLastRow(sheet: worksheet);
        }
        // Apply cleaning rules.
        private void Clean()
        {
            Range thisCell;

            // Run down the source column (skipping the header row), applying each cleaning rule.
            for (int row_offset = 1; row_offset < this.lastRow; row_offset++)
            {
                thisCell = this.sourceColumn.Offset[row_offset, 0];
                string cell_contents = thisCell.Value2;

                foreach (CleaningRule rule in config.CleaningRules)
                {
                    cell_contents = Regex.Replace(cell_contents, rule.pattern, rule.replace);
                }

                thisCell.Value2 = cell_contents;
            }
        }
        private void Extract()
        {
            Range thisCell;

            // Run down the source column (skipping the header row),
            // applying each extraction rule, stopping at the first one that matches.
            for (int row_offset = 1; row_offset < this.lastRow; row_offset++)
            {
                thisCell = this.sourceColumn.Offset[row_offset, 0];
                string cell_contents = thisCell.Value;

                foreach (ExtractRule rule in config.ExtractRules)
                {
                    Range targetRng = Utilities.TopOfNamedColumn(sheet: this.worksheet, columnName: rule.new_column);

                    if (targetRng == null)
                    {
                        targetRng = Utilities.InsertNewColumn(range: this.sourceColumn, newColumnName: rule.new_column);
                    }

                    foreach(Pattern pat in rule.Patterns)
                    {
                        Match match = Regex.Match(cell_contents, pat.pattern);

                        // Did we match?
                        if (match.Groups.Count > 1)
                        {
                            targetRng.Offset[row_offset, 0].Value = match.Groups[1].Value;
                            break;  // Don't need to search any more patterns for this rule.
                        }
                    }
                }
            }

        }
        public void Parse()
        {   
            // Apply cleaning rules.
            Clean();

            // Apply extraction rules.
            Extract();

            // Save a copy of the revised workbook.
            SaveRevised();
        }
        private void SaveRevised()
        {
            Workbook workbook = this.worksheet.Parent;
            string filename = workbook.FullName;
            string directory = System.IO.Path.GetDirectoryName(filename);
            string justTheFilename = System.IO.Path.GetFileNameWithoutExtension(filename);
            string newFilename = System.IO.Path.Combine(directory, justTheFilename + "_revised.xlsx");

            try
            {
                workbook.SaveAs(newFilename, XlSaveAsAccessMode.xlNoChange);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                newFilename = System.IO.Path.Combine(justTheFilename + "_revised");
                workbook.SaveAs(newFilename, XlSaveAsAccessMode.xlNoChange);
            }

            MessageBox.Show("Saved in '" + newFilename + "'.");
        }
        //private void SaveTemp()
        //{
        //    Workbook workbook = this.worksheet.Parent;
        //    workbook.SaveCopyAs(@"C:\tmp\inprogress.xls");
        //}
    }
}
