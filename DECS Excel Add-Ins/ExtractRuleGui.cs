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
    internal class ExtractRuleGui : RuleGui
    {
        private NotesConfig config;
        private Action<RuleGui> parentDeleteAction;
        private Action parentRuleChangedAction;
        private bool textChangedCallbackEnabled = true;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        /**
         * @brief A version of @c RuleGui that's tailored for a data extraction rule.
         */
        public ExtractRuleGui(
            int x,
            int y,
            int index,
            Panel parent,
            NotesConfig notesConfig,
            bool updateConfig = true
        )
            : base(x, y, index, parent, "extractRules")
        {
            config = notesConfig;
            base.leftTextBox.TextChanged += extractRulesDisplayNameTextBox_TextChanged;
            base.centerTextBox.TextChanged += extractRulesPatternTextBox_TextChanged;
            base.rightTextBox.LostFocus += extractRulesnewColumnTextBox_TextChanged;

            // When loading an >existing< NotesConfig object,
            // we don't want to modify the object.
            if (updateConfig)
            {
                // There needs to be an ExtractRule object (even if it's an empty placeholder) for every ExtractRuleGui object.
                config.AddExtractRule();
            }

            // The Delete button is part of the base class, but this class 'knows'
            // it's an >extract< rule that needs to be deleted.
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
        /// because WE know it's an >extract< rule.
        /// Also, because the DefineRules class creates THIS class (and not the parent RuleGui class),
        /// we'll pass the delete action along to the DefineRules class to tell it to bump the cleaning Add button upwards.
        /// </summary>
        /// <param name="ruleGui">Our parent object</param>
        
        protected void DeleteRule(RuleGui ruleGui)
        {
            config.DeleteExtractRule(index: base.index);
            parentDeleteAction(ruleGui);
        }

        /// <summary>
        /// Callback for when a @c DisplayNameTextBox object's text is changed.
        /// Insert or update Nth extract rule with this display name.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void extractRulesDisplayNameTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled)
                return;

            log.Debug("extractRulesDisplayNameTextBox_TextChanged");
            TextBox textBox = (TextBox)sender;

            // Insert or update Nth extract rule with this display Name.
            config.ChangeExtractRuleDisplayName(index: base.index, displayName: textBox.Text);
        }

        /// <summary>
        /// Callback for when a @c PatternTextBox object's text is changed.
        /// Extract the text & add to this extract rule.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void extractRulesPatternTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled)
                return;

            log.Debug("extractRulesPatternTextBox_TextChanged.");
            TextBox textBox = (TextBox)sender;
            RuleValidationResult result = Utilities.IsRegexValid(textBox.Text);

            if (result.Valid())
            {
                // Clear any previous highlighting.
                Utilities.ClearRegexInvalid(textBox);

                // Insert or update Nth extract rule with this pattern.
                config.ChangeExtractRulePattern(index: base.index, pattern: textBox.Text);

                // Alert upper-level GUI.
                parentRuleChangedAction();
            }
            else
            {
                // Highlight box to show RegEx is invalid.
                Utilities.MarkRegexInvalid(textBox: textBox, message: result.ToString());

                // Clear Nth extract rule's pattern.
                config.ChangeExtractRulePattern(index: base.index, pattern: string.Empty);
            }
        }

        /// <summary>
        /// Callback for when a @c newColumnTextBox object's text is changed.
        /// Extract the text & add to this extract rule.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void extractRulesnewColumnTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!textChangedCallbackEnabled)
                return;

            log.Debug("extractRulesnewColumnTextBox_TextChanged.");
            TextBox textBox = (TextBox)sender;

            // Insert or update Nth extract rule with this pattern.
            config.ChangeExtractRulenewColumn(index: base.index, newColumn: textBox.Text);

            // Alert upper-level GUI.
            parentRuleChangedAction();
        }

        /// <summary>
        /// Populate this GUI with a @c ExtractRule object.
        /// </summary>
        /// <param name="rule">A @c ExtractRule object to be visualized</param>
        
        public void Populate(ExtractRule rule)
        {
            if (rule == null)
                return;

            textChangedCallbackEnabled = false;
            base.leftTextBox.Text = rule.displayName;
            base.centerTextBox.Text = rule.pattern;
            base.rightTextBox.Text = rule.newColumn;
            RuleValidationResult result = Utilities.IsRegexValid(rule.pattern);

            // Validate the rule.
            if (!result.Valid())
            {
                Utilities.MarkRegexInvalid(textBox: base.centerTextBox, message: result.ToString());
            }

            // Validate the rule.
            if (string.IsNullOrEmpty(rule.newColumn))
            {
                Utilities.MarkRegexInvalid(
                    textBox: base.rightTextBox,
                    message: "newColumn is empty"
                );
            }

            textChangedCallbackEnabled = true;
        }
    }
}
