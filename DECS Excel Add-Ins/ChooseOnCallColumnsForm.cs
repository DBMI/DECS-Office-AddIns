using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace DECS_Excel_Add_Ins
{
    public partial class ChooseOnCallColumnsForm : Form
    {
        private bool haveSelectedDateColumn = false;
        private bool haveSelectedNameColumn = false;
        public string selectedDateColumn = "None";
        public string selectedNameColumn = "None";

        public ChooseOnCallColumnsForm(List<string> columnNames)
        {
            InitializeComponent();

            onCallDateColumnListBox.DataSource = null;
            onCallDateColumnListBox.Items.Clear();
            onCallDateColumnListBox.DataSource = columnNames;

            onCallNameColumnListBox.DataSource = null;
            onCallNameColumnListBox.Items.Clear();
            
            // Create a clone of the list so the ListBoxes aren't linked.
            onCallNameColumnListBox.DataSource= new List<string>(columnNames);
        }

        private void DateColumn_SelectedIndexChanged(object sender, EventArgs e)
        {
            haveSelectedDateColumn = true;
            onCallRunButton.Enabled = haveSelectedNameColumn;
        }

        private void NameColumn_SelectedIndexChanged(object sender, EventArgs e)
        {
            haveSelectedNameColumn = true;
            onCallRunButton.Enabled = haveSelectedDateColumn;
        }

        public void QuitButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        public void RunButton_Click(object sender, EventArgs e)
        {
            selectedDateColumn = onCallDateColumnListBox.SelectedItem.ToString();
            selectedDateColumn = onCallNameColumnListBox.SelectedItem.ToString();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
