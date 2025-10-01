using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public partial class MatchTextForm : Form
    {
        private Dictionary<string, Worksheet> worksheetsDict;
        private Dictionary<string, Range> redcapColumnsDict;
        private Dictionary<string, Range> artColumnsDict;

        public Range redcapIdColumn;
        public Range redcapMessageColumn;
        public Range artIdColumn;
        public Range artMessageColumn;

        public MatchTextForm()
        {
            InitializeComponent();
            LoadSheets();
        }

        private void EnableWhenReady(object sender, System.EventArgs e)
        {
            if (redcapIdColumnsListBox.SelectedItems.Count > 0 &&
                redcapMessageColumnsListBox.SelectedItems.Count > 0 &&
                artIdColumnsListBox.SelectedItems.Count > 0 &&
                artMessageColumnsListBox.SelectedItems.Count > 0)
            {
                okButton.Enabled = true;
            }
        }


        private void LoadSheets()
        {
            worksheetsDict = Utilities.GetWorksheets();
            List<string> worksheetNames = worksheetsDict.Keys.ToList<string>();
            Utilities.PopulateListBox(redcapSheetsListBox, worksheetNames, enableWhenPopulated: true);
            Utilities.PopulateListBox(artSheetsListBox, worksheetNames, enableWhenPopulated: true);
        }

        private void cancelButton_Click(object sender, System.EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void okButton_Click(object sender, System.EventArgs e)
        {
            string columnName = redcapIdColumnsListBox.SelectedItem as string;
            redcapIdColumn = redcapColumnsDict[columnName];
            columnName = redcapMessageColumnsListBox.SelectedItem as string;
            redcapMessageColumn = redcapColumnsDict[columnName];

            columnName = artIdColumnsListBox.SelectedItem as string;
            artIdColumn = artColumnsDict[columnName];
            columnName = artMessageColumnsListBox.SelectedItem as string;
            artMessageColumn = artColumnsDict[columnName];

            DialogResult = DialogResult.OK;
        }

        private void redcapSheetsListBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            // Get all the columns from this sheet & populate columns listbox.
            string selectedSheetName = redcapSheetsListBox.SelectedItem as string;
            Worksheet selectedSheet1 = worksheetsDict[selectedSheetName];

            redcapColumnsDict = Utilities.GetColumnRangeDictionary(selectedSheet1);
            List<string> columnNames = redcapColumnsDict.Keys.ToList();
            Utilities.PopulateListBox(redcapIdColumnsListBox, columnNames, enableWhenPopulated: true);
            Utilities.PopulateListBox(redcapMessageColumnsListBox, columnNames, enableWhenPopulated: true);
        }

        private void artSheetsListBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            // Get all the columns from this sheet & populate columns listbox.
            string selectedSheetName = artSheetsListBox.SelectedItem as string;
            Worksheet selectedSheet2 = worksheetsDict[selectedSheetName];

            artColumnsDict = Utilities.GetColumnRangeDictionary(selectedSheet2);
            List<string> columnNames = artColumnsDict.Keys.ToList();
            Utilities.PopulateListBox(artIdColumnsListBox, columnNames, enableWhenPopulated: true);
            Utilities.PopulateListBox(artMessageColumnsListBox, columnNames, enableWhenPopulated: true);
        }
    }
}
