using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using SimMetrics.Net;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
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
        private Dictionary<string, DateRangeColumns> selectedDateRangeColumnsDict;
        private bool disableCallbacks;
        private ColumnNamePairs dateColumnPairs;            // Such as insurance start, end dates
        private List<string> ignoredWords;
        private bool initializing;
        private List<string> selectedInfoColumns;           // Like insurer, address
        private List<string> selectedPatientDefnColumns;    // Like MRN
        private string selectedSourceSheetName;
        private Worksheet selectedSourceWorksheet;
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
            selectedDateRangeColumnsDict = new Dictionary<string, DateRangeColumns>();
            selectedInfoColumns = new List<string>();
            selectedPatientDefnColumns = new List<string>();
            ignoredWords = new List<string>() { "Date", "End", "Start" };
            disableCallbacks = false;
            initializing = true;
            List<string> sheetNames = worksheetsDict.Keys.ToList<string>();
            PopulateSourceSheets(sheetNames);
            initializing = false;
        }

        private void BuildDateRangeColumnsDict()
        {
            selectedDateRangeColumnsDict.Clear();

            foreach(ColumnNamePair pair in dateColumnPairs.GetColumnPairs())
            {
                Range col1 = availableSourceColumnsRangeDict[pair.Name1()];
                Range col2 = availableSourceColumnsRangeDict[pair.Name2()];
                string commonName = pair.CommonName();
                DateRangeColumns columnsObj = new DateRangeColumns(col1, col2, commonName);
                selectedDateRangeColumnsDict.Add(commonName, columnsObj);
            }
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
        /// Callback for when the @c dateColumnsListBox SelectedIndex is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void DateColumnsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            List<string> selectedDateColumns = new List<string>();

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = dateColumnsListBox.Items.Cast<string>().ToList();

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    selectedDateColumns.Add(thisColumn);

                    // Was the selected DATE column already in the INFO columns list?
                    // 1) Then remove it from the selected Info columns...
                    selectedInfoColumns.Remove(thisColumn);

                    // 2) ...and remove it from the Info Columns ListBox.
                    infoColumnsListBox.Items.Remove(thisColumn);

                    // Was the selected DATE column already in the PATIENT DEFINITION columns list?
                    selectedPatientDefnColumns.Remove(thisColumn);
                    patientDefinitionColumnsListBox.Items.Remove(thisColumn);
                }
                else // or DEselect it?
                {
                    selectedDateColumns.Remove(thisColumn);

                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(infoColumnsListBox, thisColumn);
                    InsertIntoListBox(patientDefinitionColumnsListBox, thisColumn);
                }
            }

            // Turn this list of selected Date columns into ColumnNamePairs object.
            dateColumnPairs = new ColumnNamePairs(selectedDateColumns, ignoredWords);

            // Remember where the pair of date columns is for each pair type ("Address", "Insurance", etc.).
            BuildDateRangeColumnsDict();

            EnableRunWhenReady();
        }

        // Checks to ensure each existing date range pair is
        // contiguous with the corresponding dates in this new row.
        private bool DateRangesContiguous(int rowOffset)
        {
            foreach (string key in selectedDateRangeColumnsDict.Keys)
            {
                DateRangeColumns colObj = selectedDateRangeColumnsDict[key];

                if (!colObj.CanMergeDates(rowOffset))
                {
                    return false;
                }
            }

            return true;
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
                    dateColumnPairs.Count() > 0 &&
                    selectedInfoColumns.Count > 0 &&
                    selectedPatientDefnColumns.Count > 0 &&
                    !string.IsNullOrEmpty(selectedSourceSheetName);
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
        /// Callback for when the @c infoColumnsListBox SelectedIndex is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void InfoColumnsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            selectedInfoColumns.Clear();
            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = infoColumnsListBox.Items.Cast<string>().ToList();

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    selectedInfoColumns.Add(thisColumn);

                    // Was the selected INFO column already in the DATE columns list?
                    // 1) Then remove it from the selected DATE columns...
                    dateColumnPairs.Remove(thisColumn);

                    // 2) ...and remove it from the Date Columns ListBox.
                    dateColumnsListBox.Items.Remove(thisColumn);

                    // Was the selected INFO column already in the PATIENT DEFINITION columns list?
                    selectedPatientDefnColumns.Remove(thisColumn);
                    patientDefinitionColumnsListBox.Items.Remove(thisColumn);
                }
                else // or DEselect it?
                {
                    selectedInfoColumns.Remove(thisColumn);

                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(dateColumnsListBox, thisColumn);
                    InsertIntoListBox(patientDefinitionColumnsListBox, thisColumn);
                }
            }

            EnableRunWhenReady();
        }

        /// <summary>
        /// Puts column back into a ListBox at the original location.
        /// </summary>
        /// <param name="listBox">ListBox object</param>
        /// <param name="columnName">str Column to insert</param>

        private void InsertIntoListBox(System.Windows.Forms.ListBox listBox, string columnName)
        {
            // Where does this column appear in the original columns list?
            List<string> availableSourceColumns = availableSourceColumnsRangeDict.Keys.ToList();
            int index = availableSourceColumns.FindIndex(c => c == columnName);

            // Only proceed if column appears in the source columns list.
            if (index >= 0)
            {
                int numInListNow = listBox.Items.Count;

                if (!listBox.Items.Contains(columnName))
                {
                    listBox.Items.Insert(Math.Min(numInListNow, index), columnName);
                }
            }
        }

        /// <summary>
        /// Overwrites DateTime object in target sheet, given named column & row number.
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

        private void OverwriteDates(Worksheet targetSheet, int rowOffset)
        {
            foreach (string key in selectedDateRangeColumnsDict.Keys)
            {
                DateRangeColumns colObj = selectedDateRangeColumnsDict[key];
                OverwriteDate(targetSheet, colObj.StartColumnName(), rowOffset, colObj.StartDate());
                OverwriteDate(targetSheet, colObj.EndColumnName(), rowOffset, colObj.EndDate());
            }

        }

        /// <summary>
        /// Callback for when the @c patientDefinitionColumnsListBox SelectedIndex is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void PatientDefinitionColumnsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            selectedPatientDefnColumns.Clear();
            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = patientDefinitionColumnsListBox.Items.Cast<string>().ToList();

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    selectedPatientDefnColumns.Add(thisColumn);

                    // Was the selected PATIENT DEFINITION column already in the DATE columns list?
                    // 1) Then remove it from the selected DATE columns...
                    dateColumnPairs.Remove(thisColumn);

                    // 2) ...and remove it from the Date Columns ListBox.
                    dateColumnsListBox.Items.Remove(thisColumn);

                    selectedInfoColumns.Remove(thisColumn);
                    infoColumnsListBox.Items.Remove(thisColumn);
                }
                else // or DEselect it?
                {
                    selectedPatientDefnColumns.Remove(thisColumn);

                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(dateColumnsListBox, thisColumn);
                    InsertIntoListBox(infoColumnsListBox, thisColumn);
                }
            }

            EnableRunWhenReady();
        }

        /// <summary>
        /// Fills the dateColumnsListBox
        /// </summary>

        private void PopulateDateColumns()
        {
            disableCallbacks = true;
            dateColumnsListBox.DataSource = null;

            List<string> availableSourceColumns = availableSourceColumnsRangeDict.Keys.ToList();
            availableSourceColumns = availableSourceColumns.Except(selectedPatientDefnColumns);
            availableSourceColumns = availableSourceColumns.Except(selectedInfoColumns);

            if (availableSourceColumns.Count > 0)
            {
                Utilities.PopulateListBox(dateColumnsListBox, availableSourceColumns);

                if (dateColumnPairs.Count() == 0)
                {
                    // Then initialize ListBox.
                    dateColumnPairs = new ColumnNamePairs(availableSourceColumns, ignoredWords);
                    dateColumnsListBox.SelectedIndex = 0;
                }
                else
                {
                    // By resetting the .DataSource property, we've wiped out any previous selections.
                    // So restore the already-selected columns.
                    foreach (string columnName in dateColumnPairs.GetColumnNames())
                    {
                        int indexToRestore = availableSourceColumns.FindIndex(r => r == columnName);

                        if (indexToRestore >= 0)
                        {
                            dateColumnsListBox.SelectedIndices.Add(indexToRestore);
                        }
                    }
                }
            }

            disableCallbacks = false;            
            EnableRunWhenReady();
        }

        /// <summary>
        /// Fills the infoColumnsListBox
        /// </summary>

        private void PopulateInfoColumns()
        {
            disableCallbacks = true;
            infoColumnsListBox.DataSource = null;

            List<string> availableSourceColumns = availableSourceColumnsRangeDict.Keys.ToList();
            availableSourceColumns = availableSourceColumns.Except(selectedPatientDefnColumns);
            List<string> dateColumns = dateColumnPairs.GetColumnNames();
            availableSourceColumns = availableSourceColumns.Except(dateColumns);

            if (availableSourceColumns.Count > 0)
            {
                Utilities.PopulateListBox(infoColumnsListBox, availableSourceColumns);

                if (selectedInfoColumns.Count == 0)
                {
                    // Then initialize ListBox.
                    selectedInfoColumns.Add(availableSourceColumns[0]);
                    infoColumnsListBox.SelectedIndex = 0;
                }
                else
                {
                    // By resetting the .DataSource property, we've wiped out any previous selections.
                    // So restore the already-selected columns.
                    foreach (string selectedColumn in selectedInfoColumns)
                    {
                        int indexToRestore = availableSourceColumns.FindIndex(r => r == selectedColumn);

                        if (indexToRestore >= 0)
                        {
                            infoColumnsListBox.SelectedIndices.Add(indexToRestore);
                        }
                    }
                }
            }

            disableCallbacks = false;
            EnableRunWhenReady();
        }

        /// <summary>
        /// Fills the patientDefinitionColumnsListBox
        /// </summary>

        private void PopulatePatientDefnColumns()
        {
            disableCallbacks = true;
            patientDefinitionColumnsListBox.DataSource = null;

            List<string> availableSourceColumns = availableSourceColumnsRangeDict.Keys.ToList();
            List<string> dateColumns = dateColumnPairs.GetColumnNames();
            availableSourceColumns = availableSourceColumns.Except(dateColumns);
            availableSourceColumns = availableSourceColumns.Except(selectedInfoColumns);

            if (availableSourceColumns.Count > 0)
            {
                Utilities.PopulateListBox(patientDefinitionColumnsListBox, availableSourceColumns);

                if (selectedPatientDefnColumns.Count == 0)
                {
                    selectedPatientDefnColumns.Add(availableSourceColumns[0]);
                    patientDefinitionColumnsListBox.SelectedIndex = 0;
                }
                else
                {
                    // Restore the already-selected columns.
                    foreach (string selectedColumn in selectedPatientDefnColumns)
                    {
                        int indexToRestore = availableSourceColumns.FindIndex(r => r == selectedColumn);

                        if (indexToRestore >= 0)
                        {
                            patientDefinitionColumnsListBox.SelectedIndices.Add(indexToRestore);
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

            // Create new worksheet.
            targetWorksheet = (Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            targetWorksheet.Name = selectedSourceSheetName + "_merged";
            int targetRowOffset = 1;

            //  & initialize with header & first row.
            // (These are offsets from first row.)
            CopyRow(0, 0);
            CopyRow(1, 1);

            // Run down the rows, looking for consecutive rows in which:
            //  ...all of the Patient Definition Columns are identical and
            //  ...all of Info Columns are close enough.
            int numRows = Utilities.FindLastRow(selectedSourceWorksheet);
            bool inGroup = false;

            Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.Application;

            for (int sourceRowOffset = 2; sourceRowOffset < numRows; sourceRowOffset++)
            {
                log.Debug("Row " + sourceRowOffset);

                if (!DateRangesContiguous(sourceRowOffset))
                {
                    // Just copy over the current line and move to the next line.
                    log.Debug("Start Date is null--skipping row.");
                    targetRowOffset++;
                    CopyRow(sourceRowOffset, targetRowOffset);
                    continue;
                }

                // If this row matches the previous row, use its dates to update the time spanned.
                if (RowsMatch(selectedPatientDefnColumns, sourceRowOffset, exact: true) &&
                    RowsMatch(selectedInfoColumns, sourceRowOffset, exact: false))
                {
                    log.Debug("In group.");

                    inGroup = true;

                    // Replace the existing start/stop dates.
                    UpdateDateRanges(sourceRowOffset);
                    OverwriteDates(targetWorksheet, targetRowOffset);
                }
                else
                {
                    log.Debug("Not in group.");
                    inGroup = false;

                    // If this row doesn't match the previous row, copy this new row to the new sheet.
                    targetRowOffset++;
                    CopyRow(sourceRowOffset, targetRowOffset);
                }

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
            List<string> availableSourceColumns = availableSourceColumnsRangeDict.Keys.ToList();
            dateColumnPairs = new ColumnNamePairs();
            PopulateDateColumns();
            PopulateInfoColumns();
            PopulatePatientDefnColumns();
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

        private void UpdateDateRanges(int rowOffset)
        {
            foreach (string key in selectedDateRangeColumnsDict.Keys)
            {
                DateRangeColumns colObj = selectedDateRangeColumnsDict[key];
                colObj.UpdateDates(rowOffset);
            }
        }
    }
}
