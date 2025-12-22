using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public partial class ChooseNumBogusRecordsForm : Form
    {
        public int numRecords;

        public ChooseNumBogusRecordsForm()
        {
            InitializeComponent();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            numRecords = (int)numRecordsUpDown.Value;
            GenerateFakeRecords recordGenerator = new GenerateFakeRecords(numRecords);
            Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}