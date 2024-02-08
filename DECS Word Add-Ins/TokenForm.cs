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
    /// Creates custom form to ask user to provide their GitLab token.
    /// </summary>
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
            System.Diagnostics.Process.Start(
                "https://docs.gitlab.com/ee/user/profile/personal_access_tokens.html"
            );
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void quitButton_Click(object sender, EventArgs e)
        {
            token = string.Empty;
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void tokenTextBox_TextChanged(object sender, EventArgs e)
        {
            token = tokenTextBox.Text;
        }
    }
}
