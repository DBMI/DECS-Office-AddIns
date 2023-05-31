using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Panel = System.Windows.Forms.Panel;
using TextBox = System.Windows.Forms.TextBox;
using log4net;

namespace DECS_Excel_Add_Ins
{
    internal class CleaningRuleGui : RuleGui
    {
        private NotesConfig config;
        private Action<RuleGui> parentDeleteAction;
        private Action parentRuleChangedAction;
        private bool textChangedCallbackEnabled = true;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public CleaningRuleGui(int x, int y, int index, Panel parent, NotesConfig notesConfig, bool updateConfig = true) : base(x, y, index, parent, "cleaningRules") 
        {
            this.config = notesConfig;
            base.leftHandTextBox.TextChanged += cleaningRulesPatternTextBox_TextChanged;
            base.rightHandTextBox.TextChanged += cleaningRulesReplaceTextBox_TextChanged;

            // When loading an >existing< NotesConfig object,
            // we don't want to modify the object.
            if (updateConfig )
            {
                // There needs to be a CleaningRule object (even if it's an empty placeholder) for every CleaningRuleGui object.
                this.config.AddCleaningRule();
            }

            // The Delete button is part of the base class, but this class 'knows'
            // it's a >cleaning< rule that needs to be deleted.
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
        // Extract the text & add to this cleaning rule.
        private void cleaningRulesPatternTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!this.textChangedCallbackEnabled) return;

            log.Debug("cleaningRulesPatternTextBox_TextChanged");
            TextBox textBox = (TextBox)sender;
            RuleValidationResult result = Utilities.IsRegexValid(textBox.Text);

            if (result.Valid())
            {
                // Clear any previous highlighting.
                Utilities.ClearRegexInvalid(textBox);
                
                // Insert or update Nth cleaning rule with this pattern.
                this.config.ChangeCleaningRulePattern(index: base.index, pattern: textBox.Text);

                // Alert upper-level GUI.
                this.parentRuleChangedAction();
            }
            else
            {
                // Highlight box to show RegEx is invalid.
                Utilities.MarkRegexInvalid(textBox: textBox, message: result.ToString());

                // Clear Nth cleaning rule's pattern.
                this.config.ChangeCleaningRulePattern(index: base.index, pattern: string.Empty);
            }
        }        
        // Extract the text & add to this cleaning rule.
        private void cleaningRulesReplaceTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled) return;
            log.Debug("cleaningRulesReplaceTextBox_TextChanged");
            TextBox textBox = (TextBox)sender;
            RuleValidationResult result = Utilities.IsRegexValid(textBox.Text);

            if (result.Valid())
            {
                // Clear any previous highlighting.
                Utilities.ClearRegexInvalid(textBox);

                // Insert or update Nth cleaning rule with this replace string.
                this.config.ChangeCleaningRuleReplace(index: base.index, replace: textBox.Text);

                // Alert upper-level GUI.
                this.parentRuleChangedAction();
            }
            else
            {
                // Highlight box to show RegEx is invalid.
                Utilities.MarkRegexInvalid(textBox: textBox, message: result.ToString());

                // Clear Nth cleaning rule's replace string.
                this.config.ChangeCleaningRuleReplace(index: base.index, replace: string.Empty);
            }
        }
        public override void Clear()
        {
            this.textChangedCallbackEnabled = false;
            base.leftHandTextBox.Text = string.Empty;
            base.rightHandTextBox.Text = string.Empty;
            this.textChangedCallbackEnabled = true;
        }
        // The RuleGui class handles the GUI stuff but this derived class needs to 'talk' to the NotesConfig structure
        // because we know it's a >cleaning< rule.
        // Also, because the DefineRules class creates THIS class (and not the parent RuleGui class),
        // we'll pass the delete action along to the DefineRules class to tell it to bump the cleaning Add button upwards.
        protected void DeleteRule(RuleGui ruleGui)
        {
            this.config.DeleteCleaningRule(index: base.index);
            this.parentDeleteAction(ruleGui);
        }
        public void Populate(CleaningRule rule)
        {
            if (rule == null) return;

            this.textChangedCallbackEnabled = false;
            base.leftHandTextBox.Text = rule.pattern;
            base.rightHandTextBox.Text = rule.replace;
            RuleValidationResult result = Utilities.IsRegexValid(rule.pattern);

            // Validate the rule.
            if (!result.Valid())
            {
                Utilities.MarkRegexInvalid(textBox: base.leftHandTextBox, message: result.ToString());
            }

            result = Utilities.IsRegexValid(rule.replace);

            if (!result.Valid())
            {
                Utilities.MarkRegexInvalid(textBox: base.rightHandTextBox, message: result.ToString());
            }

            this.textChangedCallbackEnabled = true;
        }
    }
}