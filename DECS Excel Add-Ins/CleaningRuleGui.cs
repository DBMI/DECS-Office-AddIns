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
    /**
     * @brief Specific type of @ RuleGui--representing a data cleaning rule.
     */
    internal class CleaningRuleGui : RuleGui
    {
        private NotesConfig config;
        private Action<RuleGui> parentDeleteAction;
        private Action parentRuleChangedAction;
        private bool textChangedCallbackEnabled = true;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        public CleaningRuleGui(
            int x,
            int y,
            int index,
            Panel parent,
            NotesConfig notesConfig,
            bool updateConfig = true
        )
            : base(x, y, index, parent, "cleaningRules")
        {
            config = notesConfig;
            base.leftTextBox.TextChanged += cleaningRulesDisplayNameTextBox_TextChanged;
            base.centerTextBox.TextChanged += cleaningRulesPatternTextBox_TextChanged;
            base.rightTextBox.TextChanged += cleaningRulesReplaceTextBox_TextChanged;

            // When loading an >existing< NotesConfig object,
            // we don't want to modify the object.
            if (updateConfig)
            {
                // There needs to be a CleaningRule object (even if it's an empty placeholder) for every CleaningRuleGui object.
                config.AddCleaningRule();
            }

            // The Delete button is part of the base class, but this class 'knows'
            // it's a >cleaning< rule that needs to be deleted.
            base.AssignDelete(DeleteRule);
        }

        /// <summary>
        /// Lets an external class assign this object's @c parentDeleteAction property.
        /// </summary>
        /// <param name="deleteAction">Action</param>
        public void AssignExternalDelete(Action<RuleGui> deleteAction)
        {
            parentDeleteAction = deleteAction;
        }

        /// <summary>
        /// Lets an external class assign this object's @c parentRuleChangedAction property.
        /// </summary>
        /// <param name="ruleChangedAction">Action</param>
        public void AssignExternalRuleChanged(Action ruleChangedAction)
        {
            parentRuleChangedAction = ruleChangedAction;
        }

        /// <summary>
        /// Callback for when a @c DisplayNameTextBox object's text is changed.
        /// Insert or update Nth cleaning rule with this display name.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        private void cleaningRulesDisplayNameTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled)
                return;

            log.Debug("cleaningRulesDisplayNameTextBox_TextChanged");
            TextBox textBox = (TextBox)sender;

            // Insert or update Nth cleaning rule with this display Name.
            config.ChangeCleaningRuleDisplayName(index: base.index, displayName: textBox.Text);
        }

        /// <summary>
        /// Callback for when a @c PatternTextBox object's text is changed.
        /// Extract the text & add to this cleaning rule.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        private void cleaningRulesPatternTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled)
                return;

            log.Debug("cleaningRulesPatternTextBox_TextChanged");
            TextBox textBox = (TextBox)sender;
            RuleValidationResult result = Utilities.IsRegexValid(textBox.Text);

            if (result.Valid())
            {
                // Clear any previous highlighting.
                Utilities.ClearRegexInvalid(textBox);

                // Insert or update Nth cleaning rule with this pattern.
                config.ChangeCleaningRulePattern(index: base.index, pattern: textBox.Text);

                // Alert upper-level GUI.
                parentRuleChangedAction();
            }
            else
            {
                // Highlight box to show RegEx is invalid.
                Utilities.MarkRegexInvalid(textBox: textBox, message: result.ToString());

                // Clear Nth cleaning rule's pattern.
                config.ChangeCleaningRulePattern(index: base.index, pattern: string.Empty);
            }
        }

        /// <summary>
        /// Callback for when a @c ReplaceTextBox object's text is changed.
        /// Extract the text & add to this cleaning rule.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        private void cleaningRulesReplaceTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled)
                return;
            log.Debug("cleaningRulesReplaceTextBox_TextChanged");
            TextBox textBox = (TextBox)sender;
            RuleValidationResult result = Utilities.IsRegexValid(textBox.Text);

            if (result.Valid())
            {
                // Clear any previous highlighting.
                Utilities.ClearRegexInvalid(textBox);

                // Insert or update Nth cleaning rule with this replace string.
                config.ChangeCleaningRuleReplace(index: base.index, replace: textBox.Text);

                // Alert upper-level GUI.
                parentRuleChangedAction();
            }
            else
            {
                // Highlight box to show RegEx is invalid.
                Utilities.MarkRegexInvalid(textBox: textBox, message: result.ToString());

                // Clear Nth cleaning rule's replace string.
                config.ChangeCleaningRuleReplace(index: base.index, replace: string.Empty);
            }
        }

        /// <summary>
        /// Clears the GUI.
        /// </summary>
        public override void Clear()
        {
            textChangedCallbackEnabled = false;
            base.centerTextBox.Text = string.Empty;
            base.rightTextBox.Text = string.Empty;
            textChangedCallbackEnabled = true;
        }

        /// <summary>
        /// The parent RuleGui class handles the GUI stuff but this derived class needs to 'talk' to the NotesConfig structure
        /// because WE know it's a >cleaning< rule.
        /// Also, because the DefineRules class creates THIS class (and not the parent RuleGui class),
        /// we'll pass the delete action along to the DefineRules class to tell it to bump the cleaning Add button upwards.
        /// </summary>
        /// <param name="ruleGui">Our parent object</param>
        
        protected void DeleteRule(RuleGui ruleGui)
        {
            config.DeleteCleaningRule(index: base.index);
            parentDeleteAction(ruleGui);
        }

        /// <summary>
        /// Populate this GUI with a @c CleaningRule object.
        /// </summary>
        /// <param name="rule">A @c CleaningRule object to be visualized</param>
        
        public void Populate(CleaningRule rule)
        {
            if (rule == null)
                return;

            textChangedCallbackEnabled = false;
            base.leftTextBox.Text = rule.displayName;
            base.centerTextBox.Text = rule.pattern;
            base.rightTextBox.Text = rule.replace;
            RuleValidationResult result = Utilities.IsRegexValid(rule.pattern);

            // Validate the rule.
            if (!result.Valid())
            {
                Utilities.MarkRegexInvalid(textBox: base.centerTextBox, message: result.ToString());
            }

            result = Utilities.IsRegexValid(rule.replace);

            if (!result.Valid())
            {
                Utilities.MarkRegexInvalid(textBox: base.rightTextBox, message: result.ToString());
            }

            textChangedCallbackEnabled = true;
        }
    }
}
