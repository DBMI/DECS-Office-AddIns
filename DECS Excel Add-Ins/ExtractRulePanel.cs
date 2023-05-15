using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using GroupBox = System.Windows.Forms.GroupBox;
using TextBox = System.Windows.Forms.TextBox;

namespace DECS_Excel_Add_Ins
{
    internal class ExtractRulePanel : RulePanel
    {
        private NotesConfig config;
        private Action parentDeleteAction;
        private bool textChangedCallbackEnabled = true;

        public ExtractRulePanel(int x, int y, int index, GroupBox parent, NotesConfig notesConfig, bool updateConfig = true) : base(x, y, index, parent, "extractRules")
        {
            config = notesConfig;
            leftHandTextBox.TextChanged += extractRulesPatternTextBox_TextChanged;
            rightHandTextBox.TextChanged += extractRulesnewColumnTextBox_TextChanged;

            // When loading an >existing< NotesConfig object,
            // we don't want to modify the object.
            if (updateConfig)
            {
                // There needs to be an ExtractRule object (even if it's an empty placeholder) for every ExtractRulePanel object.
                config.AddExtractRule();
            }

            // The Delete button is part of the base class, but this class 'knows'
            // it's an extract rule that needs to be deleted.
            base.AssignDelete(this.DeleteRule);
        }
        public void AssignExternalDelete(Action deleteAction)
        {
            parentDeleteAction = deleteAction;
        }
        public override void Clear()
        {
            textChangedCallbackEnabled = false;
            leftHandTextBox.Text = string.Empty;
            rightHandTextBox.Text = string.Empty;
            textChangedCallbackEnabled = true;
        }
        // The RulePanel class handles the GUI stuff but this derived class needs to 'talk' to the NotesConfig structure
        // because we know it's an >extract< rule.
        // Also, because the DefineRules class creates THIS class (and not the parent RulePanel class),
        // we'll pass the delete action along to the DefineRules class to tell it to bump the extract Add button upwards.
        protected void DeleteRule()
        {
            config.DeleteExtractRule(index: index);
            parentDeleteAction();
        }
        private void extractRulesPatternTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled) return;

            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;

            try
            {
                Regex regex = new Regex(textBox.Text);

                // Clear any previous highlighting.
                textBox.BackColor = Color.White;

                // Insert or update Nth extract rule with this pattern.
                config.ChangeExtractRulePattern(index: index, pattern: textBox.Text);
            }
            catch (ArgumentException)
            {
                // Highlight box to show RegEx is invalid.
                textBox.BackColor = Color.Pink;

                // Clear Nth cleaning rule's pattern.
                config.ChangeExtractRulePattern(index: index, pattern: string.Empty);
            }
        }
        private void extractRulesnewColumnTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled) return;

            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;

            try
            {
                Regex regex = new Regex(textBox.Text);

                // Clear any previous highlighting.
                textBox.BackColor = Color.White;

                // Insert or update Nth extract rule with this pattern.
                config.ChangeExtractRulenewColumn(index: index, newColumn: textBox.Text);
            }
            catch (ArgumentException)
            {
                // Highlight box to show RegEx is invalid.
                textBox.BackColor = Color.Pink;

                // Clear Nth cleaning rule's pattern.
                config.ChangeExtractRulenewColumn(index: index, newColumn: string.Empty);
            }
        }
        public void Populate(ExtractRule rule)
        {
            if (rule == null) return;

            textChangedCallbackEnabled = false;
            leftHandTextBox.Text = rule.pattern;
            rightHandTextBox.Text = rule.newColumn;
            textChangedCallbackEnabled = true;
        }
    }
}
