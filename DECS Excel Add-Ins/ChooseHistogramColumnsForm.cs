using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public partial class ChooseHistogramColumnsForm : Form
    {
        public string categoryColumn = string.Empty;
        public string scoreColumn = string.Empty;

        private bool allowEnable = false;
        private string dummyValue = "--None--";

        public ChooseHistogramColumnsForm(List<string> columnNames)
        {
            InitializeComponent();

            categoryColumnListBox.DataSource = null;
            categoryColumnListBox.Items.Clear();
            categoryColumnListBox.DataSource = columnNames;

            scoreColumnListBox.DataSource = null;
            scoreColumnListBox.Items.Clear();

            // Create a clone of the list so the ListBoxes aren't linked.
            List<string> scoreColumns = new List<string>(columnNames);

            // And insert a dummy value as the first item.
            scoreColumns.Insert(0, dummyValue);
            scoreColumnListBox.DataSource = scoreColumns;
           
            // NOW I'll let you turn on the "OK" button.
            allowEnable = true;
        }

        public void CancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        public void ListBoxSelect(object sender, EventArgs e)
        {
            if (allowEnable)
            {
                if (categoryColumnListBox.SelectedItems.Count > 0 && scoreColumnListBox.SelectedItems.Count > 0)
                {
                    okButton.Enabled = true;
                }
            }
        }

        public void RunButton_Click(object sender, EventArgs e)
        {
            categoryColumn = categoryColumnListBox.SelectedItem.ToString();

            try
            {
                scoreColumn = scoreColumnListBox.SelectedItem.ToString();

                if (scoreColumn == dummyValue)
                {
                    scoreColumn = string.Empty;
                }
            }
            catch (NullReferenceException)
            {
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
