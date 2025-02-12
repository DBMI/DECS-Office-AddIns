using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace DECS_Excel_Add_Ins
{
    public partial class HideThisNameForm : Form
    {
        private string preamble = @"{\rtf1\ansi ";
        public string chosenName = string.Empty;
        public string linkedName = string.Empty;
        private Func<string, List<string>> refreshSimilarNames;

        public HideThisNameForm(string nameToConsider, 
                                List<string> similarNames,
                                string leftContext, 
                                string rightContext,
                                Func<string, List<string>> _refreshSimilarNames)
        {
            InitializeComponent();
            PopulateSimilarNamesListbox(similarNames);
            contextRichTextBox.Clear();
            contextRichTextBox.Rtf = preamble + leftContext + 
                                     BoldWords(nameToConsider) +
                                     rightContext;
            linkNameButton.Enabled = false;
            refreshSimilarNames = _refreshSimilarNames;
        }

        private string BoldWords(string words)
        {
            return @"\b " + words + @"\b0 ";
        }

        public void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        public void IgnoreButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
            this.Close();
        }

        public void LinkButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        public void NewButton_Click(object sender, EventArgs e)
        {
            linkedName = string.Empty;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void PopulateSimilarNamesListbox(List<string> similarNames)
        {
            similarNamesListBox.Items.Clear();

            foreach(string similarName in similarNames) 
            { 
                similarNamesListBox.Items.Add(similarName);
            }
        }

        private void ShowAllButton_Click(Object sender, EventArgs e)
        {
            // Send an empty to "refresh" ==> send EVERYTHING.
            PopulateSimilarNamesListbox(refreshSimilarNames(string.Empty));
        }

        private void similarNamesListBox_Click(object sender, EventArgs e)
        {
            if (similarNamesListBox.SelectedItem != null)
            {
                linkedName = similarNamesListBox.SelectedItem.ToString().Trim();
                linkNameButton.Enabled = true;
            }
        }

        private void contextRichTextBox_MouseUp(object sender, MouseEventArgs e)
        {
            string rtfContents = contextRichTextBox.Text;

            if (contextRichTextBox.SelectedText == null)
            {
                // Remove all formatting.
                contextRichTextBox.Rtf = preamble + rtfContents;
            }
            else
            {
                // Highlight selected text.
                int selectionStart = contextRichTextBox.SelectionStart;
                int selectionLen = contextRichTextBox.SelectionLength;
                string leftContext = rtfContents.Substring(0, selectionStart);
                int indexOfSelectionEnd = selectionStart + selectionLen;
                int lengthOfRightContext = rtfContents.Length - indexOfSelectionEnd;
                string rightContext = rtfContents.Substring(indexOfSelectionEnd, lengthOfRightContext);
                contextRichTextBox.Rtf = preamble + leftContext +
                                     BoldWords(contextRichTextBox.SelectedText) +
                                     rightContext;

                // Find all similar names so user can link "Dr. Able Provider" with "Provider, Able, MD".
                chosenName = contextRichTextBox.SelectedText.Trim();
                linkedName = string.Empty;
                linkNameButton.Enabled = false;

                if (!string.IsNullOrEmpty(chosenName))
                {
                    // Now refresh the similar names--in case we can find a match NOW.
                    PopulateSimilarNamesListbox(refreshSimilarNames(chosenName));
                }
            }
        }
    }
}
