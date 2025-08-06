using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public partial class PlotSelectionForm : Form
    {
        public string name1Column;
        public string name2Column;
        public string sheet1Name;
        public string sheet2Name;
        public string time1Column;
        public string time2Column;
        public string value1Column;
        public string value2Column;

        private List<string> availableColumns1;
        private List<string> availableColumns2;
        private List<string> availableSheets;
        private bool disableCallbacks;
        private bool initializing;
        private Dictionary<string, Worksheet> worksheets;

        public PlotSelectionForm(Dictionary<string, Worksheet> sheets)
        {
            InitializeComponent();

            disableCallbacks = false;
            initializing = true;

            worksheets = sheets;

            List<string> sheetNames = sheets.Keys.ToList();
            PopulateSheetListboxes(sheetNames);

            initializing = false;
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void EnableRunWhenReady()
        {
            if (initializing)
            {
                return;
            }

            okButton.Enabled =
                    !string.IsNullOrEmpty(name1Column) &&
                    !string.IsNullOrEmpty(name2Column) &&
                    !string.IsNullOrEmpty(time2Column) &&
                    !string.IsNullOrEmpty(time2Column) &&
                    !string.IsNullOrEmpty(value1Column) &&
                    !string.IsNullOrEmpty(value2Column);
        }

        private void InsertIntoListBox(System.Windows.Forms.ListBox listBox, string columnName, List<string> availableNames)
        {
            // Where does this column appear in the original columns list?
            int index = availableNames.FindIndex(c => c == columnName);

            // Only proceed if column appears in the available columns list.
            if (index >= 0)
            {
                int numInListNow = listBox.Items.Count;

                if (!listBox.Items.Contains(columnName))
                {
                    listBox.Items.Insert(Math.Min(numInListNow, index), columnName);
                }
            }
        }

        private void Name1ColumnListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            disableCallbacks = true;

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();
            name1Column = string.Empty;

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--but is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    name1Column = thisColumn;
                }
                else // or DEselected?
                {
                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(timeColumn_1_ListBox, thisColumn, availableColumns1);
                    InsertIntoListBox(valueColumn_1_ListBox, thisColumn, availableColumns1);
                }
            }

            disableCallbacks = false;

            EnableRunWhenReady();
        }

        private void Name2ColumnListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            disableCallbacks = true;

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();
            name2Column = string.Empty;

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--but is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    name2Column = thisColumn;
                }
                else // or DEselected?
                {
                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(timeColumn_2_ListBox, thisColumn, availableColumns2);
                    InsertIntoListBox(valueColumn_2_ListBox, thisColumn, availableColumns2);
                }
            }

            disableCallbacks = false;

            EnableRunWhenReady();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void PopulateColumn1Listboxes(List<string> columns)
        {
            disableCallbacks = true;
            availableColumns1 = columns;
            Utilities.PopulateListBox(nameColumn_1_ListBox, columns, enableWhenPopulated: true);
            Utilities.PopulateListBox(timeColumn_1_ListBox, columns, enableWhenPopulated: true);
            Utilities.PopulateListBox(valueColumn_1_ListBox, columns, enableWhenPopulated: true);
            disableCallbacks = false;
        }

        private void PopulateColumn2Listboxes(List<string> columns)
        {
            disableCallbacks = true;
            availableColumns2 = columns;
            Utilities.PopulateListBox(nameColumn_2_ListBox, columns, enableWhenPopulated: true);
            Utilities.PopulateListBox(timeColumn_2_ListBox, columns, enableWhenPopulated: true);
            Utilities.PopulateListBox(valueColumn_2_ListBox, columns, enableWhenPopulated: true);
            disableCallbacks = false;
        }

        private void PopulateSheetListboxes(List<string> sheets)
        {
            disableCallbacks = true;
            availableSheets = sheets;
            Utilities.PopulateListBox(sheet_1_ListBox, sheets);
            Utilities.PopulateListBox(sheet_2_ListBox, sheets);
            disableCallbacks = false;
        }

        private void Sheet1ListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            disableCallbacks = true;

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();
            sheet1Name = string.Empty;

            foreach (string thisSheet in listBoxContents)
            {
                // This item is in the list box--but is it SELECTED?
                if (listBox.SelectedItems.Contains(thisSheet))
                {
                    sheet1Name = thisSheet;
                }
                else // or DEselected?
                {
                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(sheet_2_ListBox, thisSheet, availableSheets);
                }
            }

            Worksheet sheet = worksheets[sheet1Name];
            Dictionary<string, Range> headers = Utilities.GetColumnRangeDictionary(sheet);
            List<string> columnNames = headers.Keys.ToList();
            PopulateColumn1Listboxes(columnNames);
            disableCallbacks = false;

            EnableRunWhenReady();
        }

        private void Sheet2ListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            disableCallbacks = true;

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();
            sheet2Name = string.Empty;

            foreach (string thisSheet in listBoxContents)
            {
                // This item is in the list box--but is it SELECTED?
                if (listBox.SelectedItems.Contains(thisSheet))
                {
                    sheet2Name = thisSheet;
                }
                else // or DEselected?
                {
                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(sheet_1_ListBox, thisSheet, availableSheets);
                }
            }

            Worksheet sheet = worksheets[sheet2Name];
            Dictionary<string, Range> headers = Utilities.GetColumnRangeDictionary(sheet);
            List<string> columnNames = headers.Keys.ToList();
            PopulateColumn2Listboxes(columnNames);
            disableCallbacks = false;

            EnableRunWhenReady();
        }

        private void Time1ColumnListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            disableCallbacks = true;

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();
            time1Column = string.Empty;

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--but is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    time1Column = thisColumn;
                }
                else // or DEselected?
                {
                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(nameColumn_1_ListBox, thisColumn, availableColumns1);
                    InsertIntoListBox(valueColumn_1_ListBox, thisColumn, availableColumns1);
                }
            }

            disableCallbacks = false;

            EnableRunWhenReady();
        }

        private void Time2ColumnListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            disableCallbacks = true;

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();
            time2Column = string.Empty;

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--but is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    time2Column = thisColumn;
                }
                else // or DEselected?
                {
                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(nameColumn_2_ListBox, thisColumn, availableColumns2);
                    InsertIntoListBox(valueColumn_2_ListBox, thisColumn, availableColumns2);
                }
            }

            disableCallbacks = false;

            EnableRunWhenReady();
        }

        private void Value1ColumnListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            disableCallbacks = true;

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();
            value1Column = string.Empty;

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--but is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    value1Column = thisColumn;
                }
                else // or DEselected?
                {
                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(nameColumn_1_ListBox, thisColumn, availableColumns1);
                    InsertIntoListBox(timeColumn_1_ListBox, thisColumn, availableColumns1);
                }
            }

            disableCallbacks = false;

            EnableRunWhenReady();
        }

        private void Value2ColumnListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            disableCallbacks = true;

            System.Windows.Forms.ListBox listBox = sender as System.Windows.Forms.ListBox;
            List<string> listBoxContents = listBox.Items.Cast<string>().ToList();
            value2Column = string.Empty;

            foreach (string thisColumn in listBoxContents)
            {
                // This item is in the list box--but is it SELECTED?
                if (listBox.SelectedItems.Contains(thisColumn))
                {
                    value2Column = thisColumn;
                }
                else // or DEselected?
                {
                    // Since we're not using this column for dates, make it available in the other ListBoxes.
                    InsertIntoListBox(nameColumn_2_ListBox, thisColumn, availableColumns2);
                    InsertIntoListBox(timeColumn_2_ListBox, thisColumn, availableColumns2);
                }
            }

            disableCallbacks = false;

            EnableRunWhenReady();
        }
    }
}
