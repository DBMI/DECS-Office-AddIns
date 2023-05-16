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
    internal class CleaningRuleGui : RuleGui
    {
        private NotesConfig config;
        private Action<RuleGui> parentDeleteAction;
        private Action parentRuleChangedAction;
        private bool textChangedCallbackEnabled = true;

        public CleaningRuleGui(int x, int y, int index, Panel parent, NotesConfig notesConfig, bool updateConfig = true) : base(x, y, index, parent, "cleaningRules") 
        {
            config = notesConfig;
            leftHandTextBox.TextChanged += cleaningRulesPatternTextBox_TextChanged;
            rightHandTextBox.TextChanged += cleaningRulesReplaceTextBox_TextChanged;

            // When loading an >existing< NotesConfig object,
            // we don't want to modify the object.
            if (updateConfig )
            {
                // There needs to be a CleaningRule object (even if it's an empty placeholder) for every CleaningRuleGui object.
                config.AddCleaningRule();
            }

            // The Delete button is part of the base class, but this class 'knows'
            // it's a cleaning rule that needs to be deleted.
            base.AssignDelete(this.DeleteRule);
        }
        public void AssignExternalDelete(Action<RuleGui> deleteAction)
        {
            parentDeleteAction = deleteAction;
        }
        public void AssignExternalRuleChanged(Action ruleChangedAction)
        {
            parentRuleChangedAction = ruleChangedAction;
        }
        // Extract the text & add to this cleaning rule.
        private void cleaningRulesPatternTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled) return;

            TextBox textBox = (TextBox)sender;

            try
            {
                Regex regex = new Regex(textBox.Text);

                // Clear any previous highlighting.
                textBox.BackColor = Color.White;

                // Insert or update Nth cleaning rule with this pattern.
                config.ChangeCleaningRulePattern(index: index, pattern: textBox.Text);

                // Alert upper-level GUI.
                parentRuleChangedAction();
            }
            catch (ArgumentException)
            {
                // Highlight box to show RegEx is invalid.
                textBox.BackColor = Color.Pink;

                // Clear Nth cleaning rule's pattern.
                config.ChangeCleaningRulePattern(index: index, pattern: string.Empty);
            }
        }        
        // Extract the text & add to this cleaning rule.
        private void cleaningRulesReplaceTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled) return;

            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;

            try
            {
                Regex regex = new Regex(textBox.Text);

                // Clear any previous highlighting.
                textBox.BackColor = Color.White;

                // Insert or update Nth cleaning rule with this replace string.
                config.ChangeCleaningRuleReplace(index: index, replace: textBox.Text);

                // Alert upper-level GUI.
                parentRuleChangedAction();
            }
            catch (ArgumentException)
            {
                // Highlight box to show RegEx is invalid.
                textBox.BackColor = Color.Pink;

                // Clear Nth cleaning rule's replace string.
                config.ChangeCleaningRuleReplace(index: index, replace: string.Empty);
            }
        }
        public override void Clear()
        {
            textChangedCallbackEnabled = false;
            leftHandTextBox.Text = string.Empty;
            rightHandTextBox.Text = string.Empty;
            textChangedCallbackEnabled = true;
        }
        // The RuleGui class handles the GUI stuff but this derived class needs to 'talk' to the NotesConfig structure
        // because we know it's a >cleaning< rule.
        // Also, because the DefineRules class creates THIS class (and not the parent RuleGui class),
        // we'll pass the delete action along to the DefineRules class to tell it to bump the cleaning Add button upwards.
        protected void DeleteRule(RuleGui ruleGui)
        {
            config.DeleteCleaningRule(index: this.index);
            parentDeleteAction(ruleGui);
        }
        public void Populate(CleaningRule rule)
        {
            if (rule == null) return;

            textChangedCallbackEnabled = false;
            leftHandTextBox.Text = rule.pattern;
            rightHandTextBox.Text = rule.replace;
            textChangedCallbackEnabled = true;
        }
    }
}
