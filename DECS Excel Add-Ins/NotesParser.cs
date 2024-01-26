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
using Action = System.Action;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using log4net;
using System.Collections.Concurrent;
using System.Threading;
using System.Data;

namespace DECS_Excel_Add_Ins
{
    internal class NotesParser
    {
        private Application application;
        private NotesConfig config;
        private int lastCol;
        private int lastRow;
        private List<string> originalColumnNames;
        private List<string> originalSourceColumnEntries;
        private bool processAllRows;
        private ProcessingRowsSelection rowsToProcess;
        private Range sourceColumn;
        private StatusForm statusForm;
        private bool stopProcessing = false;
        private bool rulesValid;
        private Worksheet worksheet;
        private Action<ProcessingRowsSelection> worksheetChangedCallback;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        public NotesParser(Worksheet _worksheet, bool withConfigFile = true, bool allRows = true)
        {
            log.Debug("Instantiating NotesParser object.");

            application = Globals.ThisAddIn.Application;
            worksheet = _worksheet;
            worksheet.SelectionChange += WorksheetSelectionChanged;

            // Identify last row, column.
            lastCol = Utilities.FindLastCol(sheet: worksheet);
            lastRow = Utilities.FindLastRow(sheet: worksheet);

            // In development mode, we may want to run rules on just selected rows.
            processAllRows = allRows;

            rowsToProcess = WhichRowsToProcess();
            log.Debug("Processing " + rowsToProcess.NumRows().ToString() + " rows.");

            // Remember what it looked like before any additional extracted columns were added.
            originalSourceColumnEntries = new List<string>();
            originalColumnNames = Utilities.GetColumnNames(worksheet);

            if (withConfigFile)
            {
                string configFilename = NotesConfig.ChooseConfigFile();
                NotesConfig configObj = NotesConfig.ReadConfigFile(configFilename);
                UpdateConfig(configObj);
            }
        }

        internal void AssignWorksheetChangedCallback(
            Action<ProcessingRowsSelection> externalCallback
        )
        {
            worksheetChangedCallback = externalCallback;
        }

        // Apply cleaning rules.
        internal bool Clean()
        {
            log.Debug("Starting cleaning.");

            if (!HasConfig() || !rulesValid)
                return true;

            CreateStatusForm();

            log.Debug("Ordering StatusForm object .Show().");
            statusForm.Show();
            statusForm.UpdateStatusLabel("Applying cleaning rules.");

            ShowRow(1);
            RestoreOriginalSourceColumn();
            Range thisCell;

            // Run down the source column, applying each cleaning rule.
            foreach (Range row in rowsToProcess.GetRows())
            {
                if (stopProcessing)
                    return false;

                int rowNumber = row.Row;
                log.Debug("Processing row " + rowNumber.ToString());
                ShowRow(rowNumber);
                thisCell = sourceColumn.Offset[rowNumber - 1, 0];
                string cell_contents;

                try
                {
                    cell_contents = thisCell.Value2.ToString();
                }
                catch
                {
                    // There's nothing in this cell.
                    continue;
                }

                foreach (CleaningRule rule in config.ValidCleaningRules())
                {
                    try
                    {
                        cell_contents = Regex.Replace(cell_contents, rule.pattern, rule.replace);
                    }
                    catch (System.ArgumentNullException) { }

                    statusForm?.UpdateProgressBarLabel(rule.displayName ?? rule.replace);
                }

                thisCell.Value2 = cell_contents;
                statusForm?.UpdateCount();
            }

            return true;
        }

        internal void ConvertDatesToStandardFormat()
        {
            if (!HasConfig() || !config.HasDateConversionRule())
                return;

            log.Debug("Starting date conversion.");
            CreateStatusForm();

            log.Debug("Ordering StatusForm object .Show().");
            statusForm.Show();
            statusForm.UpdateStatusLabel("Applying date conversion rules.");

            Range thisCell;
            DateConverter dateConverter = new DateConverter();
            string desiredDateFormat = config.DateConversionRule.desiredDateFormat;
            statusForm?.UpdateProgressBarLabel("Converting dates to " + desiredDateFormat);

            // Run down the source column, applying each extraction rule.
            foreach (Range row in rowsToProcess.GetRows())
            {
                if (stopProcessing)
                    return;

                int rowNumber = row.Row;
                log.Debug("Processing row " + rowNumber.ToString());
                ShowRow(rowNumber);
                thisCell = sourceColumn.Offset[rowNumber - 1, 0];
                string cell_contents;

                try
                {
                    cell_contents = thisCell.Value2.ToString();
                }
                catch
                {
                    // There's nothing in this cell.
                    continue;
                }

                thisCell.Value2 = dateConverter.Convert(cell_contents, desiredDateFormat);
                statusForm?.UpdateCount();
            }

            return;
        }

        private void CreateStatusForm()
        {
            int numRows = rowsToProcess.GetRows().Count;

            if (statusForm == null || statusForm.IsDisposed)
            {
                log.Debug("Creating StatusForm object.");
                statusForm = new StatusForm(
                    _numRepetitions: numRows,
                    parentStopAction: StopProcessing
                );
            }
            else
            {
                statusForm.Reset(_numRepetitions: numRows);
            }
        }

        internal bool Extract()
        {
            log.Debug("Starting extraction.");

            if (!HasConfig() || !rulesValid)
                return true;
            CreateStatusForm();

            log.Debug("Ordering StatusForm object .Show().");
            statusForm.Show();
            statusForm.UpdateStatusLabel("Applying extraction rules.");

            Range thisCell;

            // Run down the source column, applying each extraction rule.
            foreach (Range row in rowsToProcess.GetRows())
            {
                if (stopProcessing)
                    return false;

                int rowNumber = row.Row;
                log.Debug("Processing row " + rowNumber.ToString());
                ShowRow(rowNumber);
                thisCell = sourceColumn.Offset[rowNumber - 1, 0];
                string cell_contents;

                try
                {
                    cell_contents = thisCell.Value2.ToString();
                }
                catch
                {
                    // There's nothing in this cell.
                    continue;
                }

                foreach (ExtractRule rule in config.ValidExtractRules())
                {
                    // Don't create new columns with blank names.
                    if (rule.newColumn is null || rule.newColumn.Length == 0)
                        continue;
                    log.Debug("Extracting to column '" + rule.newColumn + "'.");

                    Range targetRng = Utilities.TopOfNamedColumn(
                        sheet: worksheet,
                        columnName: rule.newColumn
                    );

                    if (targetRng == null)
                    {
                        log.Debug("Creating new column '" + rule.newColumn + "'.");
                        targetRng = Utilities.InsertNewColumn(
                            range: sourceColumn,
                            newColumnName: rule.newColumn
                        );
                        log.Debug("Created new column '" + rule.newColumn + "'.");
                    }

                    try
                    {
                        log.Debug("Attempting match using pattern '" + rule.pattern + "'.");

                        // If we get more than one extracted value, concatenate into comma-separated string.
                        List<string> extractedValues = new List<string>();

                        foreach (
                            Match match in Regex.Matches(
                                cell_contents,
                                rule.pattern,
                                RegexOptions.IgnoreCase
                            )
                        )
                        {
                            if (match.Success)
                            {
                                log.Debug("Rule matched: " + match.Value.ToString());
                                extractedValues.Add(match.Groups[1].ToString());
                            }
                        }

                        targetRng.Offset[rowNumber - 1, 0].Value += String.Join(
                            ", ",
                            extractedValues
                        );
                    }
                    catch (System.ArgumentNullException)
                    {
                        log.Error("Caught System.ArgumentNullException");
                    }

                    statusForm?.UpdateProgressBarLabel(rule.displayName ?? rule.newColumn);
                }

                statusForm?.UpdateCount();
            }

            return true;
        }

        internal bool HasConfig()
        {
            return config != null && config.SourceColumnName != string.Empty;
        }

        internal void Parse()
        {
            log.Debug("Starting parsing.");

            if (!HasConfig() || !rulesValid)
                return;

            // Apply cleaning rules.
            bool keepProcessing = Clean();

            if (!keepProcessing)
            {
                ResetAfterProcessing();
                return;
            }

            ConvertDatesToStandardFormat();

            // Apply extraction rules.
            keepProcessing = Extract();

            if (!keepProcessing)
            {
                ResetAfterProcessing();
                return;
            }

            // Save a copy of the revised workbook.
            SaveRevised();
        }

        // Allow DefineRulesForm to tell us to go back to row 1 & close the status form.
        // This is useful when we only had cleaning rules--no extract rules.
        internal void ResetAfterProcessing()
        {
            ShowRow(1);
            statusForm?.Close();
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
            int numColumns = Utilities.FindLastCol(worksheet);
            Range thisCell = (Range)worksheet.Cells[1, 1];

            // Scan along the header row and delete columns that aren't original.
            for (int col_offset = 0; col_offset < numColumns; col_offset++)
            {
                string thisColumnName = thisCell.Offset[0, col_offset].Value2.ToString();

                if (!originalColumnNames.Contains(thisColumnName))
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
            lastCol = Utilities.FindLastCol(worksheet);
        }

        private void RestoreOriginalSourceColumn()
        {
            if (!HasConfig() || !rulesValid)
                return;

            if (originalSourceColumnEntries.Count == 0)
                return;

            Range thisCell;

            // Run down the source column (skipping the header row),
            // and restore its UNcleaned contents.
            for (int row_offset = 1; row_offset < lastRow; row_offset++)
            {
                thisCell = sourceColumn.Offset[row_offset, 0];
                thisCell.Value2 = originalSourceColumnEntries[row_offset - 1];
            }
        }

        public void SaveOriginalSourceColumn()
        {
            if (!HasConfig() || !rulesValid)
                return;

            originalSourceColumnEntries = new List<string>();
            Range thisCell;

            // Run down the source column (skipping the header row),
            // and save its UNcleaned contents.
            for (int row_offset = 1; row_offset < lastRow; row_offset++)
            {
                thisCell = sourceColumn.Offset[row_offset, 0];

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

        internal void SaveRevised()
        {
            if (!HasConfig() || !rulesValid)
                return;

            statusForm.UpdateStatusLabel("Saving revised file.");
            statusForm.UpdateProgressBarLabel("Complete.");
            ResetAfterProcessing();

            // Save a copy of the revised workbook.
            Workbook workbook = worksheet.Parent;
            string filename = workbook.FullName;
            string directory = System.IO.Path.GetDirectoryName(filename);
            string justTheFilename = System.IO.Path.GetFileNameWithoutExtension(filename);
            string newFilename = System.IO.Path.Combine(
                directory,
                justTheFilename + "_extracted.xlsx"
            );

            var thread = new Thread(() =>
            {
                SaveRevised(workbook, newFilename, justTheFilename);
            });

            thread.Start();
            thread.IsBackground = true;
        }

        private void SaveRevised(Workbook workbook, string newFilename, string justTheFilename)
        {
            try
            {
                workbook.SaveCopyAs(newFilename);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                newFilename = System.IO.Path.Combine(
                    justTheFilename + "_" + Utilities.GetTimestamp()
                );
                workbook.SaveCopyAs(newFilename);
            }

            MessageBox.Show("Saved in '" + newFilename + "'.");
        }

        private void ShowRow(int row)
        {
            application.ActiveWindow.ScrollRow = row;
        }

        internal void StopProcessing()
        {
            stopProcessing = true;
        }

        // Add config structure AFTER instantiation.
        internal void UpdateConfig(NotesConfig configObj, bool updateOriginalSourceColumn = true)
        {
            // If no config already defined, we probably should offer to start defining a config file.
            if (configObj == null)
                return;

            config = configObj;

            if (config.SourceColumnName == string.Empty)
                return;

            // Find the top of the source column.
            Range rng = Utilities.TopOfNamedColumn(
                sheet: worksheet,
                columnName: config.SourceColumnName
            );

            if (rng == null)
            {
                Utilities.WarnColumnNotFound(config.SourceColumnName);
                return;
            }

            sourceColumn = rng;

            if (updateOriginalSourceColumn)
            {
                // Save the uncleaned source column.
                SaveOriginalSourceColumn();
            }

            rulesValid = ValidateAndWarn();
        }

        private bool ValidateAndWarn()
        {
            bool rulesValid = true;
            List<RuleValidationError> errors = config.ValidateRules();

            if (errors.Count > 0)
            {
                rulesValid = false;

                // Don't open duplicate forms.
                if (
                    System.Windows.Forms.Application.OpenForms.OfType<RulesErrorForm>().Count() == 0
                )
                {
                    RulesErrorForm form = new RulesErrorForm(errors);
                    form.Show();
                }
            }

            return rulesValid;
        }

        internal ProcessingRowsSelection WhichRowsToProcess()
        {
            if (processAllRows)
            {
                // This object was called in production mode ==> ALL rows.
                return new ProcessingRowsSelection(
                    Utilities.AllAvailableRows(worksheet),
                    string.Empty,
                    true
                );
            }
            else
            {
                // If the user has selected some rows, we'll process those rows only.
                Excel.Range rng = (Excel.Range)application.Selection;
                Range selectedRows = rng.Rows;
                bool allRows = false;
                string reason = string.Empty;

                if (selectedRows.Count == 1)
                {
                    try
                    {
                        int selectedRowNumber = selectedRows[0].Row;

                        if (selectedRowNumber == 0)
                        {
                            reason = "Only header row selected.";
                            selectedRows = Utilities.AllAvailableRows(worksheet);
                            allRows = true;
                            return new ProcessingRowsSelection(selectedRows, reason, allRows);
                        }
                        else if (selectedRowNumber >= lastRow) // Use equals because selectedRowNumber is zero-based.
                        {
                            reason = "Selected row outside data area.";
                            selectedRows = Utilities.AllAvailableRows(worksheet);
                            allRows = true;
                            return new ProcessingRowsSelection(selectedRows, reason, allRows);
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        reason = "No rows selected.";
                        selectedRows = Utilities.AllAvailableRows(worksheet);
                        allRows = true;
                        return new ProcessingRowsSelection(selectedRows, reason, allRows);
                    }
                }

                return new ProcessingRowsSelection(selectedRows, reason, allRows);
            }
        }

        private void WorksheetSelectionChanged(Range Target)
        {
            rowsToProcess = WhichRowsToProcess();

            if (worksheetChangedCallback != null)
            {
                worksheetChangedCallback(rowsToProcess);
            }
        }
    }
}
