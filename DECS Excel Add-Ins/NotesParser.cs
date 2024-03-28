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
    /**
     * @brief Applies rules to worksheet to clean, convert & extract required data.
     */ 
    internal class NotesParser
    {
        private Application application;
        private NotesConfig config;
        private bool haveModifiedSheet = false;
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

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_worksheet">Active worksheet</param>
        /// <param name="withConfigFile">bool: are we using a stored config file? (default = true)</param>
        /// <param name="allRows">bool: are we processing ALL the rows? (default = true)</param>
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

        /// <summary>
        /// Allows external class to assign this object's @c worksheetChangedCallback action.
        /// </summary>
        /// <param name="externalCallback">Action</param>
        
        internal void AssignWorksheetChangedCallback(
            Action<ProcessingRowsSelection> externalCallback
        )
        {
            worksheetChangedCallback = externalCallback;
        }

        /// <summary>
        /// Apply cleaning rules to the designated source column.
        /// </summary>
        /// <returns>bool</returns>
        internal bool Clean()
        {
            log.Debug("Starting cleaning.");

            if (!HasConfig() || !rulesValid)
                return true;

            CreateStatusForm();

            log.Debug("Ordering StatusForm object .Show().");
            statusForm.Show();

            if (!config.HasCleaningRules())
            {
                return true;
            }

            statusForm.UpdateStatusLabel("Applying cleaning rules.");

            ShowRow(1);

            if (haveModifiedSheet)
            {
                RestoreOriginalSourceColumn();
            }

            Range thisCell;

            // Run down the source column, applying each cleaning rule.
            foreach (Range row in rowsToProcess.GetRows())
            {
                if (stopProcessing)
                {
                    return false;
                }

                int rowNumber = row.Row;
                log.Debug("Processing row " + rowNumber.ToString());
                ShowRow(rowNumber);
                thisCell = sourceColumn.Offset[rowNumber - 1, 0];
                string cell_contents;

                try
                {
                    cell_contents = Convert.ToString(thisCell.Value2);
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
                        haveModifiedSheet = true;
                    }
                    catch (System.ArgumentNullException ex) 
                    { 
                        log.Error(ex.Message);
                    }

                    statusForm?.UpdateProgressBarLabel(rule.displayName ?? rule.replace);
                }

                try
                {
                    thisCell.Value2 = cell_contents;
                }
                catch (Exception ex)
                    when (ex is System.Runtime.InteropServices.COMException || ex is System.OutOfMemoryException)
                {
                    log.Error(ex.Message);
                }

                statusForm?.UpdateCount();
            }

            return true;
        }

        /// <summary>
        /// Applies @c DateConversion rule to convert all dates found to the desired standard format.
        /// (The idea is that this way, downstream @c ExtractRules don't have to be written to handle multiple date formats).
        /// </summary>
        
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
                {
                    return;
                }

                int rowNumber = row.Row;
                log.Debug("Processing row " + rowNumber.ToString());
                ShowRow(rowNumber);
                thisCell = sourceColumn.Offset[rowNumber - 1, 0];
                string cell_contents;

                try
                {
                    cell_contents = Convert.ToString(thisCell.Value2);
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    // There's nothing in this cell.
                    log.Error(ex.Message);
                    continue;
                }

                try
                {
                    thisCell.Value2 = dateConverter.Convert(cell_contents, desiredDateFormat);
                }
                catch (Exception ex) 
                    when (ex is System.Runtime.InteropServices.COMException || ex is System.OutOfMemoryException)
                {
                    log.Error(ex.Message);
                }

                statusForm?.UpdateCount();
            }

            return;
        }

        /// <summary>
        /// Instantiates a @c StatusForm, showing processing progress.
        /// </summary>
        
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

        /// <summary>
        /// Apply data extraction rules to the designated source column.
        /// </summary>
        /// <returns>bool</returns>
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
                {
                    return false;
                }

                int rowNumber = row.Row;
                log.Debug("Processing row " + rowNumber.ToString());
                ShowRow(rowNumber);
                thisCell = sourceColumn.Offset[rowNumber - 1, 0];
                string cell_contents;

                try
                {
                    cell_contents = Convert.ToString(thisCell.Text);
                }
                catch
                {
                    // There's nothing in this cell.
                    continue;
                }

                foreach (ExtractRule rule in config.ValidExtractRules())
                {
                    // Don't create new columns with blank names.
                    if (rule.newColumn == null || rule.newColumn.Length == 0)
                    {
                        continue;
                    }
                        
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
                            newColumnName: rule.newColumn,
                            side: config.NewColumnLocation
                        );
                        log.Debug("Created new column '" + rule.newColumn + "'.");
                    }

                    try
                    {
                        log.Debug("Attempting match using pattern '" + rule.pattern + "'.");

                        MatchCollection matches = Regex.Matches(cell_contents, rule.pattern, RegexOptions.IgnoreCase);

                        if (matches.Count > 0)
                        {
                            // If we get more than one extracted value, select the LAST one.
                            Match match = matches[matches.Count - 1];

                            if (match.Success)
                            {
                                log.Debug("Rule matched: " + Convert.ToString(match.Value));

                                // Concatenate with any existing contents.
                                string existingContents = Convert.ToString(targetRng.Offset[rowNumber - 1, 0].Value);
                                string newContents = Convert.ToString(match.Groups[1]).Trim();

                                if (existingContents != null && existingContents.Length > 0)
                                {
                                    targetRng.Offset[rowNumber - 1, 0].Value = existingContents + "; " + newContents;
                                }
                                else
                                {
                                    targetRng.Offset[rowNumber - 1, 0].Value = newContents;
                                }
                            }
                        }
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

        /// <summary>
        /// Is there a config structure already defined?
        /// </summary>
        /// <returns>bool</returns>
        internal bool HasConfig()
        {
            return config != null && config.SourceColumnName != string.Empty;
        }

        /// <summary>
        /// Main method:
        /// - Applies cleaning rules
        /// - Converts dates to standard format
        /// - Runs data extraction rules.
        /// </summary>
        
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

        /// <summary>
        /// Allow DefineRulesForm to tell us to go back to row 1 & close the status form.
        /// This is useful when we only have cleaning rules--no extract rules.
        /// </summary>
        
        internal void ResetAfterProcessing()
        {
            ShowRow(1);
            statusForm?.Close();
        }

        /// <summary>
        /// Resets Worksheet to original state, removing new columns & undoing changes to the source column.
        /// </summary>
        
        internal void ResetWorksheet()
        {
            RestoreOriginalColumns();
            RestoreOriginalSourceColumn();
        }

        /// <summary>
        /// Removes columns added during processing. 
        /// Can't just walk through the list of columns and delete the new ones,
        /// because--as they're deleted--the numbering changes.
        /// </summary>
        
        private void RestoreOriginalColumns()
        {
            bool removedColumn = false;
            int numColumns = Utilities.FindLastCol(worksheet);
            Range thisCell = (Range)worksheet.Cells[1, 1];

            // Scan along the header row and delete columns that aren't original.
            for (int col_offset = 0; col_offset < numColumns; col_offset++)
            {
                string thisColumnName = Convert.ToString(thisCell.Offset[0, col_offset].Value2);

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

        /// <summary>
        /// Undoes changes to source column.
        /// </summary>
        
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

                try
                {
                    thisCell.Value2 = originalSourceColumnEntries[row_offset - 1];
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    log.Error(ex.Message);
                }
            }

            haveModifiedSheet = false;
        }

        /// <summary>
        /// Saves the source column so we can reverse cleaning & date conversion operations.
        /// </summary>
        
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
                    originalSourceColumnEntries.Add(Convert.ToString(thisCell.Value2));
                }
            }
        }

        /// <summary>
        /// Saves the workbook as revised using a name derived from the original.
        /// </summary>
        
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
                Utilities.SaveRevised(workbook, newFilename, justTheFilename);
            });

            thread.Start();
            thread.IsBackground = true;
        }

        /// <summary>
        /// Scrolls the window to the desired row.
        /// </summary>
        /// <param name="row">int number of desired row</param>
        
        private void ShowRow(int row)
        {
            application.ActiveWindow.ScrollRow = row;
        }

        /// <summary>
        /// Sets the @c stopProcessing property to true.
        /// </summary>
        
        internal void StopProcessing()
        {
            stopProcessing = true;
        }

        /// <summary>
        /// Add config structure AFTER instantiation.
        /// </summary>
        /// <param name="configObj">@c NotesConfig object containing this set of rules</param>
        /// <param name="updateOriginalSourceColumn">bool: should we overwrite the saved copy of the source column? (default: true)</param>
        
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

        /// <summary>
        /// Run syntax validation on all rules & display errors found.
        /// </summary>
        /// <returns>bool</returns>
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

        /// <summary>
        /// Decide which rows to process.
        /// </summary>
        /// <returns>@c ProcessingRowsSelection</returns>
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

        /// <summary>
        /// Invoke the @c worksheetChangedCallback action after detecting user changed something.
        /// </summary>
        /// <param name="Target">Range that was changed</param>
        
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
