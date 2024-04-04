using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using SimMetrics.Net;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    public enum ColumnType
    {
        Date,
        Text
    }

    public partial class MergeRowsForm : Form
    {
        private Dictionary<string, Range> availableSourceColumnsRangeDict;
        private Dictionary<string, ColumnType> availableSourceColumnsTypeDict;
        private bool disableCallbacks;
        private bool initializing;
        private string candidateStartDateColumn;
        private string candidateEndDateColumn;
        private List<string> selectedCandidateColumns;  // These MIGHT be the same. If they are, we can combine their rows.
        private List<string> selectedCandidateColumnsLessDates;
        private List<string> selectedGroupColumns;      // These MUST all be the same: like insurance group, coverage start/stop dates
        private Worksheet selectedSourceWorksheet;
        private string selectedSourceSheetName;
        private Worksheet targetWorksheet;
        private Range topLeftCorner;
        private Dictionary<string, Worksheet> worksheetsDict;
        
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public MergeRowsForm()
        {
            InitializeComponent();
            worksheetsDict = Utilities.GetWorksheets();
            availableSourceColumnsRangeDict = new Dictionary<string, Range>();
            availableSourceColumnsTypeDict = new Dictionary<string, ColumnType>();
            selectedCandidateColumns = new List<string>();
            selectedCandidateColumnsLessDates = new List<string>();
            selectedGroupColumns = new List<string>();
            candidateStartDateColumn = string.Empty;
            candidateEndDateColumn = string.Empty;
            disableCallbacks = false;
            initializing = true;
            PopulateSourceSheets(worksheetsDict.Keys.ToList<string>());
            initializing = false;
        }

        /// <summary>
        /// Callback for when the @c candidateColumnsListBox SelectedIndex is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void CandidateColumnsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            candidateStartDateColumn = string.Empty;
            candidateEndDateColumn = string.Empty;
            selectedCandidateColumns.Clear();
            selectedCandidateColumnsLessDates.Clear();
            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;

            foreach (var item in listBox.SelectedItems)
            {
                string selectedColumn = item.ToString();

                if (selectedColumn.ToLower().Contains("start"))
                {
                    candidateStartDateColumn = selectedColumn;
                }
                else if (selectedColumn.ToLower().Contains("end"))
                {
                    candidateEndDateColumn = selectedColumn;
                }
                else
                {
                    selectedCandidateColumnsLessDates.Add(selectedColumn);
                }

                selectedCandidateColumns.Add(selectedColumn);

                // Was the selected CANDIDATE column already in the GROUP columns list?
                var itemToRemove = selectedGroupColumns.SingleOrDefault(r => r == selectedColumn);
                
                if (itemToRemove != null)
                    selectedGroupColumns.Remove(itemToRemove);
            }

            // A CANDIDATE column can not also be a GROUP column.
            PopulateGroupColumns();

            EnableRunWhenReady();
        }

        /// <summary>
        /// Copy row from source Worksheet to target Worksheet.
        /// </summary>
        /// <param name="sourceRowOffset">int</param>
        /// <param name="targetRowOffset">int</param>

        private void CopyRow(int sourceRowOffset, int targetRowOffset)
        {
            // Convert from offset to row number.
            int sourceRowNumber = sourceRowOffset + 1;
            int targetRowNumber = targetRowOffset + 1;

            Range sourceRange = selectedSourceWorksheet.Rows[sourceRowNumber + ":" + sourceRowNumber];
            Range targetRange = targetWorksheet.Rows[targetRowNumber + ":" + targetRowNumber];
            sourceRange.Copy(targetRange);
        }

        /// <summary>
        /// Don't let Venky push "Run" until we're ready.
        /// </summary>

        private void EnableRunWhenReady()
        {
            if (initializing)
            {
                return;
            }

            runButton.Enabled =
                    selectedCandidateColumns.Count > 0 &&
                    selectedGroupColumns.Count > 0 &&
                    !string.IsNullOrEmpty(selectedSourceSheetName) &&
                    !string.IsNullOrEmpty(candidateStartDateColumn) &&
                    !string.IsNullOrEmpty(candidateEndDateColumn);
        }

        /// <summary>
        /// Retrieves DateTime object from a named column & row number.
        /// </summary>
        /// <param name="dateColumnName">string: column name</param>
        /// <param name="rowOffset">int offset from row 1</param>

        private DateTime? GetDate(string dateColumnName, int rowOffset)
        {
            DateTime? dateTime = null;
            Range columnRange = availableSourceColumnsRangeDict[dateColumnName];

            // There's a header row but not a header column.
            int colOffset = columnRange.Column - 1;

            string currentValue;

            try
            {
                currentValue = topLeftCorner.Offset[rowOffset, colOffset].Value.ToString();
                dateTime = Utilities.ConvertExcelDate(currentValue);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }

            return dateTime;
        }

        /// <summary>
        /// Callback for when the @c groupColumnsListBox SelectedIndex is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void GroupColumnsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            selectedGroupColumns.Clear();
            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;

            foreach (var item in listBox.SelectedItems)
            {
                string selectedColumn = item.ToString();
                selectedGroupColumns.Add(selectedColumn);

                // Was the selected GROUP column already in the CANDIDATE columns list?
                var itemToRemove = selectedCandidateColumns.SingleOrDefault(r => r == selectedColumn);

                if (itemToRemove != null)
                    selectedCandidateColumns.Remove(itemToRemove);
            }

            // A GROUP column can not also be a CANDIDATE column.
            PopulateCandidateColumns();

            EnableRunWhenReady();
        }

        /// <summary>
        /// Overwrites DateTime object in target sheet, given named column & row number..
        /// </summary>
        /// <param name="targetSheet">string: name of new sheet</param>
        /// <param name="dateColumnName">string: column name</param>
        /// <param name="rowOffset">int offset from row 1</param>
        /// <param name="newTime">DateTime value to write</param>

        private void OverwriteDate(Worksheet targetSheet, string dateColumnName, int rowOffset, DateTime? newTime)
        {
            if (!newTime.HasValue)
            {
                return;
            }

            Range columnRange = availableSourceColumnsRangeDict[dateColumnName];

            // Don't write "12/31/9999" for max values.
            string newTimeString = newTime.Value.ToString();

            if (newTime == DateTime.MaxValue)
            {
                newTimeString = "";
            }

            // There's a header row but not a header column.
            int colOffset = columnRange.Column - 1;

            Range targetRange = (Range)targetSheet.Cells[1, 1];

            try
            {
                targetRange.Offset[rowOffset, colOffset].Value = newTimeString;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }
        }

        /// <summary>
        /// Fills the candidateColumnsListBox
        /// </summary>

        private void PopulateCandidateColumns()
        {
            disableCallbacks = true;

            List<string> availableSourceColumns = availableSourceColumnsRangeDict.Keys.ToList();
            availableSourceColumns = availableSourceColumns.Except(selectedGroupColumns);

            if (availableSourceColumns.Count > 0)
            {
                candidateColumnsListBox.DataSource = availableSourceColumns;

                if (selectedCandidateColumns.Count == 0)
                {
                    // Then initialize ListBox.
                    selectedCandidateColumns.Add(availableSourceColumns[0]);
                    selectedCandidateColumnsLessDates.Add(availableSourceColumns[0]);
                    candidateColumnsListBox.SelectedIndex = 0;
                }
                else
                {
                    // By resetting the .DataSource property, we've wiped out any previous selections.
                    // So restore the already-selected columns.
                    foreach (string selectedColumn in selectedCandidateColumns)
                    {
                        int indexToRestore = availableSourceColumns.FindIndex(r => r == selectedColumn);

                        if (indexToRestore >= 0)
                        {
                            candidateColumnsListBox.SelectedIndices.Add(indexToRestore);
                        }
                    }
                }
            }

            disableCallbacks = false;

            EnableRunWhenReady();
        }

        /// <summary>
        /// Fills the groupColumnsListBox
        /// </summary>

        private void PopulateGroupColumns()
        {
            disableCallbacks = true;

            List<string> availableSourceColumns = availableSourceColumnsRangeDict.Keys.ToList();
            availableSourceColumns = availableSourceColumns.Except(selectedCandidateColumns);
            groupColumnsListBox.DataSource = availableSourceColumns;

            if (availableSourceColumns.Count > 0)
            {
                if (selectedGroupColumns.Count == 0)
                {
                    // Then initialize ListBox.
                    selectedGroupColumns.Add(availableSourceColumns[0]);
                    groupColumnsListBox.SelectedIndex = 0;
                }
                else
                {
                    // By resetting the .DataSource property, we've wiped out any previous selections.
                    // So restore the already-selected columns.
                    foreach (string selectedColumn in selectedGroupColumns)
                    {
                        int indexToRestore = availableSourceColumns.FindIndex(r => r == selectedColumn);

                        if (indexToRestore >= 0)
                        {
                            groupColumnsListBox.SelectedIndices.Add(indexToRestore);
                        }
                    }
                }
            }

            disableCallbacks = false;

            EnableRunWhenReady();
        }

        /// <summary>
        /// Builds the list of sheets in the sourceSheetListBox.
        /// </summary>
        /// <param name="sheets">List of string: sheet names</param>
        private void PopulateSourceSheets(List<string> sheets)
        {
            sourceSheetListBox.DataSource = sheets;

            if (sheets.Count > 0)
            {
                selectedSourceSheetName = sheets[0];
                sourceSheetListBox.SelectedIndex = 0;
                SetupSheet();
            }
            else
            {
                selectedSourceSheetName = string.Empty;
            }
        }

        /// <summary>
        /// Callback for when the @c quitButton is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void QuitButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Determines if current & previous rows match.
        /// </summary>
        /// <param name="columnNames">List of strings: column names</param>
        /// <param name="rowOffset">int offset from row 1</param>
        /// <param name="exact">bool: do you require exact match (default) or use fuzzy match?</param>

        private bool RowsMatch(List<string> columnNames, int rowOffset, bool exact = true)
        {
            double jaroWinklerThreshold = 0.65;
            SimMetrics.Net.Metric.JaroWinkler jaroWinkler = new SimMetrics.Net.Metric.JaroWinkler();

            foreach (string columnName in columnNames)
            {
                Range columnRange = availableSourceColumnsRangeDict[columnName];
                ColumnType columnType = availableSourceColumnsTypeDict[columnName];

                // There's a header row but not a header column.
                int colOffset = columnRange.Column - 1;

                string currentValue;
                string previousValue;

                try
                {
                    previousValue = topLeftCorner.Offset[rowOffset - 1, colOffset].Value.ToString();
                    currentValue = topLeftCorner.Offset[rowOffset, colOffset].Value.ToString();

                    if (exact || columnType == ColumnType.Date)
                    {
                        if (previousValue != currentValue)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        double score = jaroWinkler.GetSimilarity(previousValue, currentValue);

                        if (score < jaroWinklerThreshold)
                        {
                            return false;
                        }
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                }
            }

            return true;
        }

        /// <summary>
        /// Callback for when the @c runButton is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void RunButton_Click(object sender, EventArgs e)
        {
            log.Debug("Starting run.");

            // Add new worksheet.
            targetWorksheet = (Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            targetWorksheet.Name = selectedSourceSheetName + "_merged";
            int targetRowOffset = 1;

            //  & initialize with header & first row.
            // (These are offsets from first row.)
            CopyRow(0, 0);
            CopyRow(1, 1);

            // Keep track of full date range spanned by Candidate columns.
            DateTime? oldStartDate = GetDate(candidateStartDateColumn, 1);
            DateTime? oldEndDate = GetDate(candidateEndDateColumn, 1);

            if (!oldStartDate.HasValue || !oldEndDate.HasValue)
            {
                MessageBox.Show("Unable to determine original group date range.");
                return;
            }

            // Run down the rows, looking for consecutive rows in which all of the Group Columns are identical
            // and the Candidate Columns are close enough.
            int numRows = Utilities.FindLastRow(selectedSourceWorksheet);
            bool inGroup = false;

            Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.Application;

            for (int sourceRowOffset = 2; sourceRowOffset < numRows; sourceRowOffset++)
            {
                log.Debug("Row " + sourceRowOffset);

                DateTime? startDate = GetDate(candidateStartDateColumn, sourceRowOffset);

                if (!startDate.HasValue)
                {
                    // Just copy over the current line and move to the next line.
                    log.Debug("Start Date is null--skipping row.");
                    targetRowOffset++;
                    CopyRow(sourceRowOffset, targetRowOffset);
                    continue;
                }

                DateTime? endDate = GetDate(candidateEndDateColumn, sourceRowOffset);

                if (!endDate.HasValue)
                {
                    // Possibly because the address is still valid.
                    endDate = DateTime.MaxValue;
                }

                // If this row matches the previous row, use its dates to update the time spanned.
                if (RowsMatch(selectedGroupColumns, sourceRowOffset, exact: true) &&
                    RowsMatch(selectedCandidateColumnsLessDates, sourceRowOffset, exact: false))
                {
                    log.Debug("In group.");

                    inGroup = true;
                    startDate = new DateTime(Math.Min(oldStartDate.Value.Ticks, startDate.Value.Ticks));
                    endDate = new DateTime(Math.Max(oldEndDate.Value.Ticks, endDate.Value.Ticks));

                    // Replace the existing start/stop dates.
                    OverwriteDate(targetWorksheet, candidateStartDateColumn, targetRowOffset, startDate);
                    OverwriteDate(targetWorksheet, candidateEndDateColumn, targetRowOffset, endDate);
                }
                else
                {
                    log.Debug("Not in group.");
                    inGroup = false;

                    // If this row doesn't match the previous row, copy this new row to the new sheet.
                    targetRowOffset++;
                    CopyRow(sourceRowOffset, targetRowOffset);
                }

                oldStartDate = startDate;
                oldEndDate = endDate;

                if (sourceRowOffset % 100 == 0)
                {
                    application.StatusBar = "Processed " + sourceRowOffset.ToString() + "/" + numRows.ToString() + " rows.";
                    System.Windows.Forms.Application.DoEvents();
                }
            }

            // If we did NOT end within a group, copy over the last row to the new sheet.
            if (!inGroup)
            {
                log.Debug("Copying final row.");
                targetRowOffset++;
                CopyRow(numRows, targetRowOffset);
            }

            log.Debug("Combining complete.");
            application.StatusBar = "Combining complete.";
        }

        /// <summary>
        /// Housekeeping for once sheet is selected.
        /// </summary>

        private void SetupSheet()
        {
            selectedSourceWorksheet = worksheetsDict[selectedSourceSheetName];
            topLeftCorner = (Range)selectedSourceWorksheet.Cells[1, 1];
            availableSourceColumnsRangeDict = Utilities.GetColumnRangeDictionary(selectedSourceWorksheet);
            availableSourceColumnsTypeDict = Utilities.GetColumnTypeDictionary(selectedSourceWorksheet);
            PopulateCandidateColumns();
            PopulateGroupColumns();
        }

        /// <summary>
        /// Callback for when the @c sourceSheetListBox SelectedIndex is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        private void SourceSheetListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing)
            {
                return;
            }

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            selectedSourceSheetName = listBox.SelectedItem.ToString();
            SetupSheet();
        }
    }
}
