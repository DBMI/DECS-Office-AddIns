using System;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public partial class NumOutputFilesForm : Form
    {
        public int? numFiles;

        public NumOutputFilesForm()
        {
            InitializeComponent();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            numFiles = null;
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            numFiles = (int)Math.Floor(numFilesUpDown.Value);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
