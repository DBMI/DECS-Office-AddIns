using System;
using System.Windows.Forms;

namespace DecsWordAddIns
{
    internal enum IcdStyle
    {
        Case,
        List
    }

    internal partial class StyleSelectionForm : Form
    {
        internal IcdStyle icdStyle = IcdStyle.Case;

        internal StyleSelectionForm()
        {
            InitializeComponent();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        private void StyleRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (caseRadioButton.Checked)
            {
                icdStyle = IcdStyle.Case;
                listRadioButton.Checked = false;
            }
            else
            {
                icdStyle = IcdStyle.List;
                caseRadioButton.Checked = false;
            }
        }
    }
}
