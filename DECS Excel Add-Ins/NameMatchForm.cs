using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public partial class NameMatchForm : Form
    {
        private string desiredName;
        public string selectedName = string.Empty;

        public NameMatchForm(string desiredName, List<string> names)
        {
            InitializeComponent();
            BuildForm(desiredName, names);
        }

        private void BuildForm(string desiredName, List<string> names)
        {
            this.desiredName = desiredName;
            nameSearchedForLabel.Text = "Searching for: " + desiredName;
            Utilities.PopulateListBox(namesListBox, names, enableWhenPopulated: true);
        }

        private void RollToName(string desiredName, List<string> names)
        {
            int indexOfFirstNameAfterDesiredName = Utilities.GetIndexOfFirstWordAfterThis(names, desiredName);
            namesListBox.TopIndex = Math.Max(0, indexOfFirstNameAfterDesiredName - 2);
            namesListBox.SelectedIndex = indexOfFirstNameAfterDesiredName - 1;
        }

        private void cancelButton_Click(object sender, System.EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void okButton_Click(object sender, System.EventArgs e)
        {
            selectedName = namesListBox.SelectedItem.ToString();
            DialogResult = DialogResult.OK;
        }

        private void NameMatchForm_Load(object sender, EventArgs e)
        {
            List<string> names = Utilities.GetListBoxContents(namesListBox);
            RollToName(desiredName, names);
        }
    }
}
