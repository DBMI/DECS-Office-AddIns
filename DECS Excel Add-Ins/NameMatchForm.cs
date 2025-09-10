using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public partial class NameMatchForm : Form
    {
        private string desiredName;
        private List<string> possibleNames;
        private List<string> names;

        public string selectedName = string.Empty;
        public bool quit = false;


        public NameMatchForm(string desiredName, List<string> names, List<string> possibleNames = null)
        {
            InitializeComponent();

            this.desiredName = desiredName;
            this.names = names;
            this.possibleNames = possibleNames;
            BuildForm();
        }

        private void BuildForm()
        {
            nameSearchedForLabel.Text = "Searching for: " + desiredName;
            Utilities.PopulateListBox(namesListBox, names, enableWhenPopulated: true);
        }

        private void RollToName()
        {
            if (possibleNames.Count == 0)
            {
                int indexOfFirstNameAfterDesiredName = Utilities.GetIndexOfFirstWordAfterThis(names, desiredName);
                namesListBox.TopIndex = Math.Max(0, indexOfFirstNameAfterDesiredName - 2);
                namesListBox.SelectedIndex = indexOfFirstNameAfterDesiredName - 1;
            }
            else
            {
                int indexOfFirstName = names.IndexOf(possibleNames[0]);
                namesListBox.TopIndex = Math.Max(0, indexOfFirstName - 1);

                int indexOfLastName = names.IndexOf(possibleNames.Last());

                for (int i = indexOfFirstName; i <= indexOfLastName; i++)
                {
                    namesListBox.SetSelected(i, true);
                }

                // Don't let user select multiple names.
                okButton.Enabled = indexOfLastName == indexOfFirstName;
            }
        }

        private void cancelButton_Click(object sender, System.EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void NameMatchForm_Load(object sender, EventArgs e)
        {
            RollToName();
        }

        private void okButton_Click(object sender, System.EventArgs e)
        {
            selectedName = namesListBox.SelectedItem.ToString();
            DialogResult = DialogResult.OK;
        }

        private void quitButton_Click(object sender, EventArgs e)
        {
            quit = true;
            DialogResult = DialogResult.Abort;
        }

        // Only let user OK the match if just ONE name is selected.
        private void namesListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (namesListBox.SelectedItems.Count == 1)
            {
                okButton.Enabled = true;
            }
            else
            {
                okButton.Enabled = false;
            }
        }
    }
}
