using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DECS_Excel_Add_Ins
{
    public partial class MatchSetupForm : Form
    {
        private Dictionary<string, Worksheet> worksheetsDict;
        private Dictionary<string, Range> sourceColumnsDict;
        private Dictionary<string, Range> targetColumnsDict;

        public Range idColumn;
        public Range sourceColumn;
        public Range targetColumn;

        public MatchSetupForm()
        {
            InitializeComponent();
            LoadSheets();
        }

        private void EnableWhenReady(object sender, System.EventArgs e)
        {
            if (sourceNameColumnListBox.SelectedItems.Count > 0 &&
                idColumnListBox.SelectedItems.Count > 0 &&
                targetNameColumnListBox.SelectedItems.Count > 0)
            {
                okButton.Enabled = true;
            }
        }

        private void LoadSheets()
        {
            worksheetsDict = Utilities.GetWorksheets();
            List<string> worksheetNames = worksheetsDict.Keys.ToList<string>();
            Utilities.PopulateListBox(sourceSheetsListBox, worksheetNames, enableWhenPopulated: true);
            Utilities.PopulateListBox(targetSheetsListBox, worksheetNames, enableWhenPopulated: true);
        }

        private void sourceSheetsListBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            // Get all the columns from this sheet & populate columns listbox.
            string selectedSourceSheetName = sourceSheetsListBox.SelectedItem as string;
            Worksheet selectedSourceSheet = worksheetsDict[selectedSourceSheetName];

            sourceColumnsDict = Utilities.GetColumnRangeDictionary(selectedSourceSheet);
            List<string> columnNames = sourceColumnsDict.Keys.ToList();
            Utilities.PopulateListBox(sourceNameColumnListBox, columnNames, enableWhenPopulated: true);
            Utilities.PopulateListBox(idColumnListBox, columnNames, enableWhenPopulated: true);
        }

        private void targetSheetsListBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            // Get all the columns from this sheet & populate columns listbox.
            string selectedTargetSheetName = targetSheetsListBox.SelectedItem as string;
            Worksheet selectedTargetSheet = worksheetsDict[selectedTargetSheetName];

            targetColumnsDict = Utilities.GetColumnRangeDictionary(selectedTargetSheet);
            List<string> columnNames = targetColumnsDict.Keys.ToList();
            Utilities.PopulateListBox(targetNameColumnListBox, columnNames, enableWhenPopulated: true);
        }

        private void okButton_Click(object sender, System.EventArgs e)
        {
            string sourceColumnName = sourceNameColumnListBox.SelectedItem as string;
            sourceColumn = sourceColumnsDict[sourceColumnName];
            string targetColumnName = targetNameColumnListBox.SelectedItem as string;
            targetColumn = targetColumnsDict[targetColumnName];
            string idColumnName = idColumnListBox.SelectedItem as string;
            idColumn = sourceColumnsDict[idColumnName];

            DialogResult = DialogResult.OK;
        }

        private void cancelButton_Click(object sender, System.EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void helpButton_Click(object sender, System.EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/DBMI/DECS-Office-AddIns/blob/main/DECS%20Excel%20Add-Ins/help%20files/MatchPhysicians/MatchPhysicians.md");

        }
    }
}
