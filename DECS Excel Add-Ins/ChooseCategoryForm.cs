using Newtonsoft.Json.Bson;
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
    public partial class ChooseCategoryForm : Form
    {
        public string selectedCategory = "None";

        public ChooseCategoryForm(List<string> columnNames)
        {
            InitializeComponent();

            columnNamesListBox.DataSource = null;
            columnNamesListBox.Items.Clear();
            columnNamesListBox.DataSource = columnNames;
        }

        private void ColumnNamesListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            runButton.Enabled = true;
        }

        public void QuitButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        public void RunButton_Click(object sender, EventArgs e)
        {
            selectedCategory = columnNamesListBox.SelectedItem.ToString();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }   
    }
}
