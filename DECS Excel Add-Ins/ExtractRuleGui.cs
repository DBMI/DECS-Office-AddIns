using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Panel = System.Windows.Forms.Panel;
using TextBox = System.Windows.Forms.TextBox;

namespace DECS_Excel_Add_Ins
{
    internal class ExtractRuleGui : RuleGui
    {
        private NotesConfig config;
        private Action<RuleGui> parentDeleteAction;
        private Action parentRuleChangedAction;
        private bool textChangedCallbackEnabled = true;

        public ExtractRuleGui(int x, int y, int index, Panel parent, NotesConfig notesConfig, bool updateConfig = true) : base(x, y, index, parent, "extractRules")
        {
            this.config = notesConfig;
            base.leftHandTextBox.TextChanged += extractRulesPatternTextBox_TextChanged;
            base.rightHandTextBox.LostFocus += extractRulesnewColumnTextBox_TextChanged;

            // When loading an >existing< NotesConfig object,
            // we don't want to modify the object.
            if (updateConfig)
            {
                // There needs to be an ExtractRule object (even if it's an empty placeholder) for every ExtractRuleGui object.
                this.config.AddExtractRule();
            }

            // The Delete button is part of the base class, but this class 'knows'
            // it's an extract rule that needs to be deleted.
            base.AssignDelete(this.DeleteRule);
        }
        public void AssignExternalDelete(Action<RuleGui> deleteAction)
        {
            this.parentDeleteAction = deleteAction;
        }
        public void AssignExternalRuleChanged(Action ruleChangedAction)
        {
            this.parentRuleChangedAction = ruleChangedAction;
        }
        public override void Clear()
        {
            this.textChangedCallbackEnabled = false;
            base.leftHandTextBox.Text = string.Empty;
            base.rightHandTextBox.Text = string.Empty;
            this.textChangedCallbackEnabled = true;
        }
        // The RuleGui class handles the GUI stuff but this derived class needs to 'talk' to the NotesConfig structure
        // because we know it's an >extract< rule.
        // Also, because the DefineRules class creates THIS class (and not the parent RuleGui class),
        // we'll pass the delete action along to the DefineRules class to tell it to bump the extract Add button upwards.
        protected void DeleteRule(RuleGui ruleGui)
        {
            this.config.DeleteExtractRule(index: base.index);
            this.parentDeleteAction(ruleGui);
        }
        private void extractRulesPatternTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!this.textChangedCallbackEnabled) return;

            TextBox textBox = (TextBox)sender;

            try
            {
                Regex regex = new Regex(textBox.Text);

                // Clear any previous highlighting.
                textBox.BackColor = Color.White;

                // Insert or update Nth extract rule with this pattern.
                this.config.ChangeExtractRulePattern(index: base.index, pattern: textBox.Text);

                // Alert upper-level GUI.
                this.parentRuleChangedAction();
            }
            catch (ArgumentException)
            {
                // Highlight box to show RegEx is invalid.
                textBox.BackColor = Color.Pink;

                // Clear Nth cleaning rule's pattern.
                this.config.ChangeExtractRulePattern(index: base.index, pattern: string.Empty);
            }
        }
        private void extractRulesnewColumnTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!this.textChangedCallbackEnabled) return;

            TextBox textBox = (TextBox)sender;

            try
            {
                Regex regex = new Regex(textBox.Text);

                // Clear any previous highlighting.
                textBox.BackColor = Color.White;

                // Insert or update Nth extract rule with this pattern.
                this.config.ChangeExtractRulenewColumn(index: base.index, newColumn: textBox.Text);

                // Alert upper-level GUI.
                this.parentRuleChangedAction();
            }
            catch (ArgumentException)
            {
                // Highlight box to show RegEx is invalid.
                textBox.BackColor = Color.Pink;

                // Clear Nth cleaning rule's pattern.
                this.config.ChangeExtractRulenewColumn(index: base.index, newColumn: string.Empty);
            }
        }
        public void Populate(ExtractRule rule)
        {
            if (rule == null) return;

            this.textChangedCallbackEnabled = false;
            base.leftHandTextBox.Text = rule.pattern;
            base.rightHandTextBox.Text = rule.newColumn;
            this.textChangedCallbackEnabled = true;
        }
    }
}