using System;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public partial class UseCalforniaOrAllUsaForm : Form
    {
        public SviScope scope = SviScope.Unknown;

        public UseCalforniaOrAllUsaForm()
        {
            InitializeComponent();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            if (radioButtonCalifornia.Checked)
            {
                scope = SviScope.California;
            }

            if (radioButtonUSA.Checked)
            {
                scope = SviScope.USA;
            }

            DialogResult = DialogResult.OK;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void radioButtonCalifornia_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonCalifornia.Checked)
            {
                radioButtonUSA.Checked = false;
            }
        }

        private void radioButtonUSA_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonUSA.Checked)
            {
                radioButtonCalifornia.Checked = false;
            }
        }
    }
}
