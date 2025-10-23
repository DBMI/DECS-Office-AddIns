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
    public partial class SelectMetricsForm : Form
    {
        public List<string> selectedMetrics;

        public SelectMetricsForm(List<string> metrics)
        {
            InitializeComponent();
            Utilities.PopulateListBox(metricsListBox, metrics);
            selectedMetrics = new List<string>();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            selectedMetrics.Clear();
            selectedMetrics = metricsListBox.SelectedItems.Cast<string>().ToList();
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
