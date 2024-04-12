using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace DECS_Excel_Add_Ins
{
    public partial class MergeNotesForm : Form
    {
        private Dictionary<string, Range> availableSourceColumnsDict;
        private bool disableCallbacks;
        private bool initializing;
        private string selectedDateColumn;
        private string selectedIndexColumn;
        private List<string> selectedSourceColumns;
        private string selectedSourceSheetName;
        private string selectedTargetSheetName;
        private Dictionary<string, Worksheet> worksheetsDict;

        public MergeNotesForm()
        {
            InitializeComponent();
            worksheetsDict = Utilities.GetWorksheets();
            availableSourceColumnsDict = new Dictionary<string, Range>();
            selectedSourceColumns = new List<string>();
            disableCallbacks = false;
            initializing = true;
            PopulateSourceSheets(worksheetsDict.Keys.ToList<string>());
            PopulateTargetSheets(worksheetsDict.Keys.ToList<string>());
            initializing = false;
        }

        // Given all the rows where a particular source index (like CSN) is found,
        // select the one where there is the most data across all the source columns.
        // In case of a tie, select the most recent (if date column is specified.)
        private Range BestRow(List<Range> rows)
        {
            Range bestRange = null;
            int mostValuesPresent = 0;
            DateTime mostRecentDate = DateTime.MinValue;

            foreach(Range row in rows)
            {
                List<string> values = DataValues(row);
                int numNonEmpties = Utilities.NumElementsPresent(values);

                if (numNonEmpties > mostValuesPresent)
                {
                    mostValuesPresent = numNonEmpties;
                    bestRange = row;
                }
                else if (numNonEmpties == mostValuesPresent)
                {
                    // Check the date.
                    DateTime thisDate = DateValue(row);

                    if (thisDate > mostRecentDate)
                    {
                        mostRecentDate = thisDate;
                        bestRange = row;
                    }
                }
            }

            return bestRange;
        }

        // Get the source values--for the provided source column on this row.
        private string DataValue(Range sourceRowRange, string sourceColumnName)
        {
            Range sourceColRange = availableSourceColumnsDict[sourceColumnName];
            Range dataRange = Utilities.ThisRowThisColumn(sourceRowRange, sourceColRange);

            try
            {
                return Convert.ToString(dataRange.Value2);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return string.Empty;
            }
        }

        // Get all the source values--one for each selected source column on this row.
        private List<string> DataValues(Range sourceRow)
        {
            List<string> data = new List<string>();

            foreach(string sourceColumnName in selectedSourceColumns)
            {
                string thisValue = DataValue(sourceRow, sourceColumnName);
                data.Add(thisValue);
            }

            return data;
        }

        // Get the date value for this row.
        private DateTime DateValue(Range sourceRowRange)
        {
            Range dateColumnRange = availableSourceColumnsDict[selectedDateColumn];
            Range dataRange = Utilities.ThisRowThisColumn(sourceRowRange, dateColumnRange);

            try
            {
                if (DateTime.TryParse(dataRange.Value2, out DateTime result))
                {
                    return result;
                }
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }

            return DateTime.MinValue;
        }

        private void DateColumnListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    selectedDateColumn = thisColumn;

                    // Was the selected DATE column already the INDEX column?
                    if (selectedIndexColumn == thisColumn)
                    {
                        // 1) Then remove it as the selected Index...
                        selectedIndexColumn = string.Empty;
                    }

                    // 2) ...and remove it from the Index Columns ListBox.
                    indexColumnListBox.Items.Remove(thisColumn);

                    // Was the selected DATE column already one of the SOURCE columns?
                    // 1) Then remove it from the selected Source columns...
                    selectedSourceColumns.Remove(thisColumn);

                    // 2) ...and remove it from the Index Columns ListBox.
                    sourceColumnsListBox.Items.Remove(thisColumn);
                }
                else // or DEselected?
                {
                    selectedDateColumn = string.Empty;

                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(indexColumnListBox, thisColumn);
                    InsertIntoListBox(sourceColumnsListBox, thisColumn);
                }
            }

            EnableRunWhenReady();
        }

        private void EnableRunWhenReady()
        {
            if (initializing)
            {
                return;
            }

            runButton.Enabled = 
                    selectedSourceColumns.Count > 0 &&
                    !string.IsNullOrEmpty(selectedDateColumn) &&
                    !string.IsNullOrEmpty(selectedIndexColumn) &&
                    !string.IsNullOrEmpty(selectedTargetSheetName);
        }

        private Dictionary<string, List<Range>> GetIndexValues(string worksheetName)
        {
            Worksheet selectedWorksheet = worksheetsDict[worksheetName];
            int lastRowNumber = Utilities.FindLastRow(selectedWorksheet);
            Dictionary<string, Range> columnsDict = Utilities.GetColumnRangeDictionary(selectedWorksheet);
            Range indexColumn = columnsDict[selectedIndexColumn];
            int indexColumnNumber = indexColumn.Column;
            Dictionary<string, List<Range>> indices = new Dictionary<string, List<Range>>();
            Range indexPosition;
            string indexValue;

            // Start at row 2 to skip header row.
            for (int rowNumber = 2; rowNumber <= lastRowNumber; rowNumber++)
            {
                indexPosition = (Range)selectedWorksheet.Cells[rowNumber, indexColumnNumber];
                indexValue = Convert.ToString(indexPosition.Value2);

                if (indices.ContainsKey(indexValue))
                {
                    List<Range> ranges = indices[indexValue];
                    ranges.Add(indexPosition);
                    indices[indexValue] = ranges;
                }
                else
                {
                    indices.Add(indexValue, new List<Range>() { indexPosition });
                }
            }

            return indices;
        }

        private bool IndexColumnListBoxContains(string columnName)
        {
            List<string> columnsOffered = indexColumnListBox.DataSource as List<string>;

            if (columnsOffered != null)
            {
                return columnsOffered.Contains(columnName);
            }

            return false;
        }

        private bool IndexColumnListBoxContains(List<string> list)
        {
            List<string> columnsOffered = indexColumnListBox.DataSource as List<string>;

            if (columnsOffered != null)
            {
                foreach (string column in columnsOffered)
                {
                    if (list.Contains(column)) return true;
                }
            }

            return false;
        }
        /// <summary>
        /// Callback for when the @c indexColumnListBox SelectedIndex is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void IndexColumnListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    selectedIndexColumn = thisColumn;

                    // Was the selected INDEX column already the DATE column?
                    if (selectedDateColumn == thisColumn)
                    {
                        // 1) Then remove it as the selected Index...
                        selectedDateColumn = string.Empty;
                    }

                    // 2) ...and remove it from the Date Columns ListBox.
                    dateColumnListBox.Items.Remove(thisColumn);

                    // Was the selected INDEX column already one of the SOURCE columns?
                    // 1) Then remove it from the selected Source columns...
                    selectedSourceColumns.Remove(thisColumn);

                    // 2) ...and remove it from the SOURCE Columns ListBox.
                    sourceColumnsListBox.Items.Remove(thisColumn);
                }
                else // or DEselected?
                {
                    selectedIndexColumn = string.Empty;

                    // Since we're not using this column for index, make it available in the other ListBoxes.
                    InsertIntoListBox(dateColumnListBox, thisColumn);
                    InsertIntoListBox(sourceColumnsListBox, thisColumn);
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
            List<string> availableSourceColumns = availableSourceColumnsDict.Keys.ToList();
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
        private void PopulateDateColumn()
        {
            disableCallbacks = true;
            List<string> availableSourceColumns = availableSourceColumnsDict.Keys.ToList();
            availableSourceColumns = availableSourceColumns.Except(selectedSourceColumns);
            availableSourceColumns = availableSourceColumns.Except(selectedIndexColumn);

            // Need to prepend "None" to column list in case there IS no date column.
            availableSourceColumns = availableSourceColumns.Prepend("None").ToList();
            Utilities.PopulateListBox(dateColumnListBox, availableSourceColumns);

            // Is there a "Date" column?
            var tagged = availableSourceColumns.Select((item, i) => new { Item = item, Index = (int?)i });
            int? index = (from pair in tagged
                          where pair.Item.ToUpper().Contains("DATE")
                          select pair.Index).FirstOrDefault();

            if (index.HasValue)
            {
                dateColumnListBox.SelectedIndex = index.Value;
                selectedDateColumn = availableSourceColumns[index.Value];
            }
            else
            {
                // Then default to "None".
                dateColumnListBox.SelectedIndex = 0;
                selectedDateColumn = availableSourceColumns[0];
            }

            disableCallbacks = false;
            EnableRunWhenReady();
        }

        private void PopulateIndexColumns()
        {
            disableCallbacks = true;

            List<string> availableSourceColumns = availableSourceColumnsDict.Keys.ToList();
            availableSourceColumns = availableSourceColumns.Except(selectedSourceColumns);
            availableSourceColumns = availableSourceColumns.Except(selectedDateColumn);

            Utilities.PopulateListBox(indexColumnListBox, availableSourceColumns);

            if (availableSourceColumns.Count > 0)
            {
                selectedIndexColumn = availableSourceColumns[0];
                indexColumnListBox.SelectedIndex = 0;
            }
            else
            {
                selectedIndexColumn = string.Empty;
            }

            disableCallbacks = false;
            EnableRunWhenReady();
        }

        private void PopulateSourceColumns()
        {
            disableCallbacks = true;

            List<string> availableSourceColumns = availableSourceColumnsDict.Keys.ToList();
            availableSourceColumns = availableSourceColumns.Except(selectedIndexColumn);
            availableSourceColumns = availableSourceColumns.Except(selectedDateColumn);
            Utilities.PopulateListBox(sourceColumnsListBox, availableSourceColumns);

            selectedSourceColumns.Clear();

            if (availableSourceColumns.Count > 0)
            {
                selectedSourceColumns.Add(availableSourceColumns[0]);
                sourceColumnsListBox.SelectedIndex = 0;
            }

            disableCallbacks = false;
            EnableRunWhenReady();
        }

        private void PopulateSourceSheets(List<string> sheets)
        {
            sourceSheetListBox.DataSource = sheets;

            if (sheets.Count > 0)
            {
                selectedSourceSheetName = sheets[0];
                sourceSheetListBox.SelectedIndex = 0;
                Worksheet selectedSheet = worksheetsDict[selectedSourceSheetName];
                availableSourceColumnsDict = Utilities.GetColumnRangeDictionary(selectedSheet);
                PopulateDateColumn();
                PopulateIndexColumns();
                PopulateSourceColumns();
            }
            else
            {
                selectedSourceSheetName = string.Empty;
            }
        }

        private void PopulateTargetSheets(List<string> sheets)
        {
            targetSheetListBox.DataSource = sheets;

            if (sheets.Count > 1)
            {
                selectedTargetSheetName = sheets[1];
                targetSheetListBox.SelectedIndex = 1;
                PopulateIndexColumns();
            }
            else
            {
                selectedTargetSheetName = string.Empty;
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
        /// Callback for when the @c runButton is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void RunButton_Click(object sender, EventArgs e)
        {
            // Get all the values in the SOURCE sheet's index column.
            Dictionary<string, List<Range>> sourceIndices = GetIndexValues(worksheetName: selectedSourceSheetName);

            // Get all the values in the TARGET sheet's index column.
            Dictionary<string, List<Range>> targetIndices = GetIndexValues(worksheetName: selectedTargetSheetName);

            Worksheet targetWorksheet = worksheetsDict[selectedTargetSheetName];
            Range targetA1 = (Range)targetWorksheet.Cells[1, 1];
            Dictionary<string, Range> targetColumns = new Dictionary<string, Range>();

            // Create new column with same name in TARGET sheet.
            foreach (string sourceColumnName in selectedSourceColumns)
            {
                // Create new column with same name in TARGET sheet.
                Range targetColumnRange = Utilities.InsertNewColumn(range: targetA1, newColumnName: sourceColumnName);
                targetColumns.Add(sourceColumnName, targetColumnRange);
            }

            Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.Application;
            int valueIndex = 0;
            int numIndices = sourceIndices.Count;
            
            // For each SOURCE index value:
            foreach (KeyValuePair<string, List<Range>> entry in sourceIndices)
            {
                // Find the best row.
                Range bestRow = BestRow(entry.Value);

                if (bestRow != null)
                {
                    foreach (string sourceColumnName in selectedSourceColumns)
                    {
                        string thisValue = DataValue(sourceRowRange: bestRow, sourceColumnName: sourceColumnName);
                        List<Range> targetRowRanges;

                        // Find the row(s) in the TARGET sheet.
                        try
                        {
                            targetRowRanges = targetIndices[entry.Key];
                        }
                        // Then this index value from the source sheet isn't present in the target sheet.
                        catch (System.Collections.Generic.KeyNotFoundException) 
                        { 
                            continue; 
                        }

                        // Find the column in the TARGET sheet.
                        Range targetColumnRange = targetColumns[sourceColumnName];

                        Range targetValueRange;

                        // Find the intersection(s).
                        foreach(Range tgtRow in targetRowRanges)
                        {
                            targetValueRange = Utilities.ThisRowThisColumn(rowRange: tgtRow, columnRange: targetColumnRange);
                            
                            // Insert the value from the best row into the TARGET sheet.
                            targetValueRange.Value2 = thisValue;
                        }
                    }
                }

                valueIndex++;

                if (valueIndex % 100 == 0)
                {
                    application.StatusBar = "Processed " + valueIndex.ToString() + "/" + numIndices.ToString() + " rows.";
                }
            }

            application.StatusBar = "Merge complete.";
            SaveRevised();
        }

        internal void SaveRevised()
        {
            // Save a copy of the revised workbook.
            Workbook workbook = (Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            string filename = workbook.FullName;
            string directory = System.IO.Path.GetDirectoryName(filename);
            string justTheFilename = System.IO.Path.GetFileNameWithoutExtension(filename);
            string newFilename = System.IO.Path.Combine(
                directory,
                justTheFilename + "_merged.xlsx"
            );

            var thread = new Thread(() =>
            {
                Utilities.SaveRevised(workbook, newFilename, justTheFilename);
            });

            thread.Start();
            thread.IsBackground = true;
        }

        /// <summary>
        /// Callback for when the @c sourceColumnsListBox SelectedIndex is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void SourceColumnsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            selectedSourceColumns = new List<string>();
            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    selectedSourceColumns.Add(thisColumn);

                    // Was the selected SOURCE column already the INDEX column?
                    if (selectedIndexColumn == thisColumn)
                    {
                        // 1) Then remove it as the selected Index...
                        selectedIndexColumn = string.Empty;
                    }

                    // 2) ...and remove it from the Index Columns ListBox.
                    indexColumnListBox.Items.Remove(thisColumn);

                    // Was the selected SOURCE column already the DATE columns?
                    if (selectedDateColumn == thisColumn)
                    {
                        // 1) Then remove it from the selected Date columns...
                        selectedDateColumn = string.Empty;
                    }

                    // 2) ...and remove it from the DATE Columns ListBox.
                    dateColumnListBox.Items.Remove(thisColumn);
                }
                else // or DEselected?
                {
                    selectedSourceColumns.Remove(thisColumn);

                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(dateColumnListBox, thisColumn);
                    InsertIntoListBox(indexColumnListBox, thisColumn);
                }
            }
                
            EnableRunWhenReady();
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
            Worksheet selectedSheet = worksheetsDict[selectedSourceSheetName];
            availableSourceColumnsDict = Utilities.GetColumnRangeDictionary(selectedSheet);
            PopulateDateColumn();
            PopulateIndexColumns();
            PopulateSourceColumns();
        }

        /// <summary>
        /// Callback for when the @c targetSheetListBox SelectedIndex is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void TargetSheetListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing)
            {
                return;
            }

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            selectedTargetSheetName = listBox.SelectedItem.ToString();
        }
    }
}
