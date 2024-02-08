using System;
using System.Windows.Forms;

namespace DecsWordAddIns
{
    /// <summary>
    /// Enumeration to capture which output style we want.
    /// </summary>
    internal enum IcdStyle
    {
        Case,
        List
    }

    /// <summary>
    /// Custom form to ask user which output style is desired for ICD list.
    /// </summary>
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
