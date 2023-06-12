using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DecsWordAddIns
{
    internal partial class DeliveryTypeForm : Form
    {
        internal DeliveryType deliveryType;

        internal DeliveryTypeForm()
        {
            InitializeComponent();
        }

        private void oneDriveRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton button = (RadioButton)sender;
            vrdRadioButton.Checked = !button.Checked;

            if (button.Checked)
            {
                deliveryType = DeliveryType.OneDrive;
            }
        }

        private void vrdRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton button = (RadioButton)sender;
            oneDriveRadioButton.Checked = !button.Checked;

            if (button.Checked)
            {
                deliveryType = DeliveryType.VRD;
            }
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
