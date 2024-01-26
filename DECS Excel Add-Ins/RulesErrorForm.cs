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
    public partial class RulesErrorForm : Form
    {
        public RulesErrorForm(List<RuleValidationError> errorList)
        {
            InitializeComponent();
            rulesErrorFormLabel.Text =
                Environment.NewLine + String.Join(Environment.NewLine, errorList);
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}
