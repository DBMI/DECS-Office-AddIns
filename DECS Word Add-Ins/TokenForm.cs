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
    public partial class TokenForm : Form
    {
        public string token;

        public TokenForm()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel1.LinkVisited = true;
            System.Diagnostics.Process.Start("https://docs.gitlab.com/ee/user/profile/personal_access_tokens.html");
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void quitButton_Click(object sender, EventArgs e)
        {
            this.token = string.Empty;
            this.DialogResult= DialogResult.Cancel;
            this.Close();
        }

        private void tokenTextBox_TextChanged(object sender, EventArgs e)
        {
            this.token = tokenTextBox.Text;
        }
    }
}