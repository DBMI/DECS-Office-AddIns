using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace DECS_Excel_Add_Ins
{
    internal class NotesParser
    {
        private NotesConfig config;
        private int lastCol;
        private int lastRow;
        private List<string> originalColumnNames;
        private List<string> originalSourceColumnEntries;
        private Range sourceColumn;
        private Worksheet worksheet;
        private StatusForm statusForm;

        public NotesParser(Worksheet worksheet, bool withConfigFile = true)
        {
            this.worksheet = worksheet;

            // Identify last row, column.
            this.lastCol = Utilities.FindLastCol(sheet: worksheet);
            this.lastRow = Utilities.FindLastRow(sheet: worksheet);

            // Remember what it looked like before any additional extracted columns were added.
            this.originalColumnNames = Utilities.GetColumnNames(worksheet);

            if (withConfigFile)
            {
                string configFilename = NotesConfig.ChooseConfigFile();
                NotesConfig configObj = NotesConfig.ReadConfigFile(configFilename);
                UpdateConfig(configObj);
            }
        }
        // Apply cleaning rules.
        internal void Clean(BackgroundWorker bw = null)
        {
            if (!HasConfig()) return;

            RestoreOriginalSourceColumn();
            Range thisCell;
            int progressPercentage = 0;

            // Run down the source column (skipping the header row), applying each cleaning rule.
            for (int row_offset = 1; row_offset < this.lastRow; row_offset++)
            {
                thisCell = this.sourceColumn.Offset[row_offset, 0];
                string cell_contents = thisCell.Value2.ToString();

                foreach (CleaningRule rule in config.CleaningRules)
                {
                    try
                    {
                        cell_contents = Regex.Replace(cell_contents, rule.pattern, rule.replace);
                    }
                    catch (System.ArgumentNullException)
                    {
                    }

                    bw?.ReportProgress(progressPercentage, rule.replace);
                    statusForm?.UpdateProgressBarLabel(rule.replace);
                }

                thisCell.Value2 = cell_contents;

                // Do this only if bw is not null.
                if (this.lastRow > 1)
                {
                    progressPercentage = 100 * row_offset / (this.lastRow - 1);
                }
                else
                {
                    progressPercentage = 100;
                }

                statusForm?.UpdateProgressBar(progressPercentage);
            }
        }
        internal void Extract(BackgroundWorker bw = null)
        {
            if (!HasConfig()) return;

            Range thisCell;
            int progressPercentage = 0;

            // Run down the source column (skipping the header row),
            // applying each extraction rule, stopping at the first one that matches.
            for (int row_offset = 1; row_offset < this.lastRow; row_offset++)
            {
                thisCell = this.sourceColumn.Offset[row_offset, 0];
                string cell_contents = thisCell.Value.ToString();

                foreach (ExtractRule rule in config.ExtractRules)
                {
                    // Don't create new columns with blank names.
                    if (rule.newColumn is null || rule.newColumn.Length == 0) continue;

                    Range targetRng = Utilities.TopOfNamedColumn(sheet: this.worksheet, columnName: rule.newColumn);

                    if (targetRng == null)
                    {
                        targetRng = Utilities.InsertnewColumn(range: this.sourceColumn, newColumnName: rule.newColumn);
                    }

                    try
                    {
                        Match match = Regex.Match(cell_contents, rule.pattern);

                        // Did we match?
                        if (match.Groups.Count > 1)
                        {
                            targetRng.Offset[row_offset, 0].Value = match.Groups[1].Value;
                        }
                    }
                    catch (System.ArgumentNullException)
                    {
                    }

                    bw?.ReportProgress(progressPercentage, rule.newColumn);
                    statusForm?.UpdateProgressBarLabel(rule.newColumn);
                }

                // Do this only if bw is not null.
                if (this.lastRow > 1)
                {
                    progressPercentage = (100 * row_offset / (this.lastRow - 1));
                }
                else
                {
                    progressPercentage = 100;
                }

                statusForm?.UpdateProgressBar(progressPercentage);
            }
        }
        internal bool HasConfig()
        {
            return config != null && 
                config.SourceColumn != string.Empty;
        }
        internal void Parse()
        {
            if (!HasConfig()) return;

            this.statusForm = new StatusForm();
            this.statusForm.Show();

            // Apply cleaning rules.
            Clean();

            // Apply extraction rules.
            Extract();

            this.statusForm.Close();

            // Save a copy of the revised workbook.
            SaveRevised();

        }
        internal void ResetWorksheet()
        {
            RestoreOriginalColumns();
            RestoreOriginalSourceColumn();
        }
        // Can't just walk through the list of columns and delete the new ones, 
        // because--as they're deleted--the numbering changes.
        private void RestoreOriginalColumns()
        {
            bool removedColumn = false;
            int numColumns = Utilities.FindLastCol(this.worksheet);
            Range thisCell = (Range)this.worksheet.Cells[1, 1];

            // Scan along the header row and delete columns that aren't original.
            for (int col_offset = 0; col_offset < numColumns; col_offset++)
            {
                string thisColumnName = thisCell.Offset[0, col_offset].Value2.ToString();

                if (!this.originalColumnNames.Contains(thisColumnName))
                {
                    thisCell.Offset[0, col_offset].EntireColumn.Delete();
                    removedColumn = true;
                    break;
                }
            }

            if (removedColumn)
            {
                // Repeat until we there are no more added columns found.
                RestoreOriginalColumns();
            }

            // Need to refresh this count, now that we may have deleted a column.
            this.lastCol = Utilities.FindLastCol(this.worksheet);
        }
        private void RestoreOriginalSourceColumn()
        {
            if (!HasConfig()) return;

            Range thisCell;

            // Run down the source column (skipping the header row),
            // and restore its UNcleaned contents.
            for (int row_offset = 1; row_offset < this.lastRow; row_offset++)
            {
                thisCell = this.sourceColumn.Offset[row_offset, 0];
                thisCell.Value2 = originalSourceColumnEntries[row_offset - 1];
            }
        }
        public void SaveOriginalSourceColumn()
        {
            if (!HasConfig()) return;

            originalSourceColumnEntries = new List<string>();
            Range thisCell;

            // Run down the source column (skipping the header row),
            // and save its UNcleaned contents.
            for (int row_offset = 1; row_offset < this.lastRow; row_offset++)
            {
                thisCell = this.sourceColumn.Offset[row_offset, 0];

                if (thisCell.Value2 == null)
                {
                    originalSourceColumnEntries.Add(string.Empty);
                }
                else
                {
                    originalSourceColumnEntries.Add(thisCell.Value2.ToString());
                }
            }
        }
        private void SaveRevised()
        {
            if (!HasConfig()) return;

            Workbook workbook = this.worksheet.Parent;
            string filename = workbook.FullName;
            string directory = System.IO.Path.GetDirectoryName(filename);
            string justTheFilename = System.IO.Path.GetFileNameWithoutExtension(filename);
            string newFilename = System.IO.Path.Combine(directory, justTheFilename + "_revised.xlsx");

            try
            {
                workbook.SaveCopyAs(newFilename);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                newFilename = System.IO.Path.Combine(justTheFilename + "_" + Utilities.GetTimestamp());
                workbook.SaveCopyAs(newFilename);
            }

            MessageBox.Show("Saved in '" + newFilename + "'.");
        }
        // Add config structure AFTER instantiation.
        internal void UpdateConfig(NotesConfig configObj, bool updateOriginalSourceColumn = true)
        {
            // If no config already defined, we probably should offer to start defining a config file.
            if (configObj == null) { return; }

            this.config = configObj;

            if (this.config.SourceColumn == string.Empty) return;

            // Find the top of the source column.
            Range rng = Utilities.TopOfNamedColumn(sheet: worksheet, columnName: this.config.SourceColumn);

            if (rng == null)
            {
                Utilities.WarnColumnNotFound(this.config.SourceColumn);
                return;
            }

            this.sourceColumn = rng;

            if (updateOriginalSourceColumn)
            {
                // Save the uncleaned source column.
                SaveOriginalSourceColumn();
            }
        }
    }
}
