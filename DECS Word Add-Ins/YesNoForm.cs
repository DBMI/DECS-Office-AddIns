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
    /// <summary>
    /// Creates custom form to ask user if there's a SlicerDicer file to be parsed.
    /// </summary>
    public partial class YesNoForm : Form
    {
        public bool fileExists = false;

        public YesNoForm()
        {
            InitializeComponent();
        }

        private void yesButton_Click(object sender, EventArgs e)
        {
            this.fileExists = true;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void noButton_Click(object sender, EventArgs e)
        {
            this.fileExists = false;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void quitButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
