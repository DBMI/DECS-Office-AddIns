using System;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public partial class ChooseAOrBForm : Form
    {
        private string _optionA;
        private string _optionB;
        public string choice = string.Empty;

        public ChooseAOrBForm(string headline, string optionA, string optionB)
        {
            InitializeComponent();
            headlineLabel.Text = headline;
            _optionA = optionA;
            _optionB = optionB;
            radioButtonA.Text = optionA;
            radioButtonB.Text = optionB;
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            if (radioButtonA.Checked)
            {
                choice = _optionA;
            }

            if (radioButtonB.Checked)
            {
                choice = _optionB;
            }

            DialogResult = DialogResult.OK;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void radioButtonA_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonA.Checked)
            {
                radioButtonB.Checked = false;
            }
        }

        private void radioButtonB_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonB.Checked)
            {
                radioButtonA.Checked = false;
            }
        }
    }
}
