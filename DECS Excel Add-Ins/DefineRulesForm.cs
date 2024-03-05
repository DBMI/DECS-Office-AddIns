using Microsoft.Office.Core;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Button = System.Windows.Forms.Button;
using Excel = Microsoft.Office.Interop.Excel;
using log4net;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Form to allow user to define a text-parsing rule.
     */
    internal partial class DefineRulesForm : Form
    {
        private NotesConfig config = new NotesConfig();
        private string configFilename = string.Empty;

        private const int BUTTON_HEIGHT = 30;
        private readonly Font BUTTON_FONT = new Font("Microsoft San Serif", 14.25f, FontStyle.Bold);
        private const int BUTTON_WIDTH = 40;
        private const int BUTTON_X = 1275;
        private readonly int BUTTON_Y_OFFSET = (int)(RuleGui.Height() - BUTTON_HEIGHT) / 2;

        private const int PANEL_X = 5;
        private const int PANEL_Y = 10;
        private readonly int Y_STEP = RuleGui.Height();

        private List<RuleGui> cleaningRules = new List<RuleGui>();
        private List<RuleGui> extractRules = new List<RuleGui>();

        private NotesParser parser;

        private bool configLoading = false;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        internal DefineRulesForm(NotesParser _parser)
        {
            log.Debug("Instantiating DefineRulesForm");
            parser = _parser;
            InitializeComponent();
            PopulateDateFormatsListBox();
            PopulateSourceColumnListBox();
            AddCleaningRule();
            AddExtractRule();
            parser.AssignWorksheetChangedCallback(ShowSelectedRows);
            ShowSelectedRows();
            SetRunButtonStatus();
        }

        /// <summary>
        /// Add a @c CleaningRule.
        /// </summary>
        /// <param name="rule">CleaningRule object</param>
        /// <param name="updateConfig">Bool: Should we update the stored config file? (default: true)</param>
        
        private void AddCleaningRule(CleaningRule rule = null, bool updateConfig = true)
        {
            log.Debug("Adding cleaning rule.");
            // How many do we have now? (Index is zero-based.)
            int nextIndex = NumRulesThisType(parent: cleaningRulesPanel);
            int panelY = PANEL_Y + (Y_STEP * nextIndex);

            CleaningRuleGui cleaningRuleGui = new CleaningRuleGui(
                x: PANEL_X,
                y: panelY,
                index: nextIndex,
                parent: cleaningRulesPanel,
                notesConfig: config,
                updateConfig: updateConfig
            );

            // Tell the cleaning rule panel to let us know when ITS parent (RuleGui)'s
            // Delete button or Enable checkbox is pressed.
            cleaningRuleGui.AssignExternalDelete(DeleteCleaningRule);
            cleaningRuleGui.AssignDisable(DisableCleaningRule);
            cleaningRuleGui.AssignEnable(EnableCleaningRule);

            // Have the cleaning rule panel tell us when text changes so
            // we can invoke the ShowCleaningResult method.
            cleaningRuleGui.AssignExternalRuleChanged(RegisterChanges);
            cleaningRuleGui.Populate(rule);

            // Keep a list.
            cleaningRules.Add(cleaningRuleGui);

            // Line everything up.
            RearrangeCleaningControls();

            // Should we turn on the Run button?
            SetRunButtonStatus();
        }

        /// <summary>
        /// Add an @c ExtractRule.
        /// </summary>
        /// <param name="rule">ExtractRule object</param>
        /// <param name="updateConfig">Bool: Should we update the stored config file? (default: true)</param>
        
        private void AddExtractRule(ExtractRule rule = null, bool updateConfig = true)
        {
            log.Debug("Adding extraction rule.");
            // How many do we have now? (Index is zero-based.)
            int nextIndex = NumRulesThisType(parent: extractRulesPanel);
            int panelY = PANEL_Y + (Y_STEP * nextIndex);

            ExtractRuleGui extractRuleGui = new ExtractRuleGui(
                x: PANEL_X,
                y: panelY,
                index: nextIndex,
                parent: extractRulesPanel,
                notesConfig: config,
                updateConfig: updateConfig
            );

            // Tell the extract rule panel to let us know when ITS parent (RuleGui)'s
            // Delete button or Enable checkbox is pressed.
            extractRuleGui.AssignExternalDelete(DeleteExtractRule);
            extractRuleGui.AssignDisable(DisableExtractRule);
            extractRuleGui.AssignEnable(EnableExtractRule);

            // Have the extract rule panel tell us when text changes so
            // we can invoke the ShowExtractResult method.
            extractRuleGui.AssignExternalRuleChanged(RegisterChanges);
            extractRuleGui.Populate(rule);

            // Keep a list.
            extractRules.Add(extractRuleGui);

            // Line everything up.
            RearrangeExtractControls();

            // Should we turn on the Run button?
            SetRunButtonStatus();
        }

        /// <summary>
        /// List<ulong>all the rules in this GUI.
        /// </summary>
        /// <returns>List<RuleGui></returns>
        private List<RuleGui> AllRules()
        {
            List<RuleGui> rules = CleaningRules();
            rules.AddRange(ExtractRules());
            return rules;
        }

        private bool CheckExtractRulesUseful()
        {
            List<ExtractRule> rules = config.ExtractRules;

            // Alert user about enabled extract rules that have no capture groups.
            foreach (ExtractRule rule in rules)
            {
                if (rule.enabled && rule.pattern != null)
                {
                    if (!Utilities.HasCaptureGroups(rule.pattern))
                    {
                        string message = "Rule '" + rule.displayName + "' does not extract any data. Run anyway?";
                        string title = "Rule doesn't do anything.";
                        MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                        DialogResult result = MessageBox.Show(
                            message,
                            title,
                            buttons,
                            MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button1
                        );

                        if (result == DialogResult.No)
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// List<ulong>all the Cleaning rules in this GUI.
        /// </summary>
        /// <returns>List<RuleGui></returns>
        // https://stackoverflow.com/a/12127025/18749636
        private List<RuleGui> CleaningRules()
        {
            List<Panel> cleaningRuleGuis = cleaningRulesPanel.Controls.OfType<Panel>().ToList();

            // Assemble the list of RuleGui objects to which these Panels belong.
            List<RuleGui> rules = cleaningRuleGuis.Select(o => (RuleGui)o.Tag).ToList();
            return rules;
        }

        /// <summary>
        /// Callback for when the @c cleaningRulesAddButton is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void cleaningRulesAddButton_Click(object sender, EventArgs e)
        {
            AddCleaningRule();
        }

        /// <summary>
        /// Callback for when the @c clearButton is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void clearButton_Click(object sender, EventArgs e)
        {
            DeleteAllRules();

            // Wipe out the accumulated object.
            config = new NotesConfig();
            parser.UpdateConfig(config);

            // Restore any cleaning/extraction done.
            parser.ResetWorksheet();

            // Should we turn off the Run button?
            SetRunButtonStatus();
        }

        /// <summary>
        /// Callback for when the @c dateConversionEnabledCheckBox is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void dateConversionEnabledCheckBox_Click(object sender, EventArgs e)
        {
            CheckBox checkBox = sender as CheckBox;
            dateFormatsListBox.Enabled = checkBox.Checked;
            config.DateConversionRule.enabled = checkBox.Checked;
            SetRunButtonStatus();
        }

        /// <summary>
        /// Callback for when the @c dateFormatsListBox SelectedIndex is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void dateFormatsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListBox listBox = sender as ListBox;
            string selectedDateFormat = listBox.SelectedItem.ToString();
            config.DateConversionRule.desiredDateFormat = selectedDateFormat;
        }

        /// <summary>
        /// Deletes all the rules in the GUI.
        /// </summary>
        
        private void DeleteAllRules()
        {
            List<RuleGui> rules = AllRules();

            // Start with the highest indices & delete in descending order.
            // (So we don't try to delete objects which have already changed their index.)
            rules.OrderByDescending(x => x.Index()).ToList().ForEach(r => r.Delete());
        }

        /// <summary>
        /// Deletes one @c CleaningRule from the GUI.
        /// </summary>
        /// <param name="cleaningRuleGui">CleaningRule that's a part of this rule set.</param>
        
        internal void DeleteCleaningRule(RuleGui cleaningRuleGui)
        {
            // Remove this RuleGui from the controls.
            cleaningRulesPanel.Controls.Remove(cleaningRuleGui.PanelObj);

            // Remove from our list of CleaningRuleGui objects.
            cleaningRules.RemoveAll(r => r == cleaningRuleGui);

            cleaningRuleGui.PanelObj.Dispose();

            RearrangeCleaningControls();

            // Should we turn off the Run button?
            SetRunButtonStatus();
        }

        /// <summary>
        /// Deletes one @c ExtractRule from the GUI.
        /// </summary>
        /// <param name="extractRuleGui">ExtractRule that's a part of this rule set.</param>
        
        internal void DeleteExtractRule(RuleGui extractRuleGui)
        {
            // Remove this RuleGui from the controls.
            extractRulesPanel.Controls.Remove(extractRuleGui.PanelObj);

            // Remove from our list of CleaningRuleGui objects.
            extractRules.RemoveAll(r => r == extractRuleGui);

            extractRuleGui.PanelObj.Dispose();

            RearrangeExtractControls();

            // Should we turn off the Run button?
            SetRunButtonStatus();
        }

        /// <summary>
        /// Disables one @c CleaningRule from the GUI.
        /// </summary>
        /// <param name="ruleGui">CleaningRule that's a part of this rule set.</param>
        
        internal void DisableCleaningRule(RuleGui ruleGui)
        {
            log.Debug("Disabling cleaning rule " + ruleGui.Index().ToString() + ".");

            config.DisableCleaningRule(ruleGui.Index());

            // Should we turn off the Run button?
            SetRunButtonStatus();
        }

        /// <summary>
        /// Disables one @c ExtractRule from the GUI.
        /// </summary>
        /// <param name="ruleGui">ExtractRule that's a part of this rule set.</param>
        
        internal void DisableExtractRule(RuleGui ruleGui)
        {
            log.Debug("Disabling extract rule " + ruleGui.Index().ToString() + ".");

            config.DisableExtractRule(ruleGui.Index());

            // Should we turn off the Run button?
            SetRunButtonStatus();
        }

        /// <summary>
        /// Callback for when the @c discardButton is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void discardButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Enables one @c CleaningRule from the GUI.
        /// </summary>
        /// <param name="ruleGui">CleaningRule that's a part of this rule set.</param>
        
        internal void EnableCleaningRule(RuleGui ruleGui)
        {
            log.Debug("Enabling cleaning rule " + ruleGui.Index().ToString() + ".");

            config.EnableCleaningRule(ruleGui.Index());

            // Should we turn on the Run button?
            SetRunButtonStatus();
        }

        /// <summary>
        /// Enables one @c ExtractRule from the GUI.
        /// </summary>
        /// <param name="ruleGui">ExtractRule that's a part of this rule set.</param>
        
        internal void EnableExtractRule(RuleGui ruleGui)
        {
            log.Debug("Enabling extract rule " + ruleGui.Index().ToString() + ".");

            config.EnableExtractRule(ruleGui.Index());

            // Should we turn on the Run button?
            SetRunButtonStatus();
        }

        /// <summary>
        /// Assembles the list of RuleGui objects in the GUI.
        /// </summary>
        /// <returns>List<RuleGui></returns>
        private List<RuleGui> ExtractRules()
        {
            List<Panel> extractRuleGuis = extractRulesPanel.Controls.OfType<Panel>().ToList();

            // Assemble the list of RuleGui objects to which these Panels belong.
            List<RuleGui> rules = extractRuleGuis.Select(o => (RuleGui)o.Tag).ToList();
            return rules;
        }

        /// <summary>
        /// Callback for when the @c extractRulesAddButton is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void extractRulesAddButton_Click(object sender, EventArgs e)
        {
            AddExtractRule();
        }

        private void LeftRadioButton_Click(object sender, EventArgs e)
        {
            if (configLoading)
                return;

            // Change the left/right preference.
            config.NewColumnLocation = InsertSide.Left;
            parser.UpdateConfig(config);
        }

        /// <summary>
        /// Callback for when the @c loadToolStripMenuItem is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            log.Debug("Loading config file.");

            configFilename = NotesConfig.ChooseConfigFile();
            NotesConfig configLoaded = NotesConfig.ReadConfigFile(configFilename);

            if (configLoaded != null)
            {
                // Temporarily disable all other callbacks until we're done loading.
                configLoading = true;

                // Initialize the NotesConfig object.
                config = configLoaded;

                // Link it to the Parser.
                parser.UpdateConfig(config);

                // Clear what's already here.
                DeleteAllRules();

                // Create & populate boxes, but without changing the config variable,
                // because that's already set.
                foreach (CleaningRule rule in config.CleaningRules)
                {
                    AddCleaningRule(rule: rule, updateConfig: false);
                }

                PopulateDateConversionRule();

                foreach (ExtractRule rule in config.ExtractRules)
                {
                    AddExtractRule(rule: rule, updateConfig: false);
                }

                // Show preferred location for new column.
                SetNewColumnLocation(config.NewColumnLocation);

                // Roll the source column listbox to the proper column name.
                SetSourceColumn(config.SourceColumnName);
            }

            // Should we turn on the Run button?
            SetRunButtonStatus();

            // Regard all further alarms.
            configLoading = false;

            // Don't automatically run. The user might have wanted to open for development.
        }

        /// <summary>
        /// How many rules now in the GUI?
        /// </summary>
        /// <param name="parent">Parent @c Panel object</param>
        /// <returns>int</returns>
        private int NumRulesThisType(Panel parent)
        {
            List<Panel> panels = parent.Controls.OfType<Panel>().ToList();
            List<RuleGui> rulesThisType = panels.Select(o => (RuleGui)o.Tag).ToList();
            return rulesThisType.Count;
        }

        /// <summary>
        /// Populates the @ DateConversionRule from the stored config file.
        /// </summary>
        
        private void PopulateDateConversionRule()
        {
            dateConversionEnabledCheckBox.Checked = config.DateConversionRule.enabled;
            string loadedDateFormat = config.DateConversionRule.desiredDateFormat;

            if (string.IsNullOrEmpty(loadedDateFormat))
                return;

            if (dateFormatsListBox.Items.Contains(loadedDateFormat))
            {
                dateFormatsListBox.SelectedItem = loadedDateFormat;
            }
            else
            {
                string message = "Format '" + loadedDateFormat + "' is not supported.";
                string title = "Unsupported date format";
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                DialogResult result = MessageBox.Show(
                    message,
                    title,
                    buttons,
                    MessageBoxIcon.Warning
                );

                if (result == DialogResult.OK)
                {
                    return;
                }
            }
        }

        /// <summary>
        /// Populates the @c DateFormatsListBox from all date formats supported by the @c DateConverter object.
        /// </summary>
        
        private void PopulateDateFormatsListBox()
        {
            dateFormatsListBox.DataSource = null;
            dateFormatsListBox.Items.Clear();
            DateConverter converter = new DateConverter();
            dateFormatsListBox.DataSource = converter.SupportedDateFormats();
        }

        /// <summary>
        /// Populates the @c SourceColumnListBox using all available column names.
        /// </summary>
        
        private void PopulateSourceColumnListBox()
        {
            sourceColumnListBox.DataSource = null;
            sourceColumnListBox.Items.Clear();
            sourceColumnListBox.DataSource = Utilities.GetColumnNames((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
        }

        /// <summary>
        /// Rearranges the RuleGui objects now in the GUI in response to request to add a new rule.
        /// </summary>
        /// <param name="rules">All the RuleGui objects now present</param>
        /// <param name="addButton">The addButton just pressed</param>
        
        private void RearrangeControls(List<RuleGui> rules, Button addButton)
        {
            int panelY = PANEL_Y + Y_STEP;

            if (rules != null && rules.Count > 0)
            {
                foreach (RuleGui ruleGui in rules)
                {
                    if (ruleGui != null)
                    {
                        panelY = PANEL_Y + (Y_STEP * ruleGui.Index());
                        ruleGui.ResetLocation(x: PANEL_X, y: panelY);
                    }
                }
            }

            panelY = PANEL_Y + (Y_STEP * rules.Count);
            Point addButtonPosit = new Point(BUTTON_X, panelY + BUTTON_Y_OFFSET);
            addButton.Location = addButtonPosit;
        }

        /// <summary>
        /// Rearranges the CleaningControls now in the GUI in response to request to add a new rule.
        /// </summary>
        
        internal void RearrangeCleaningControls()
        {
            RearrangeControls(rules: cleaningRules, addButton: cleaningRulesAddButton);
        }

        /// <summary>
        /// Rearranges the ExtractControls now in the GUI in response to request to add a new rule.
        /// </summary>
        
        internal void RearrangeExtractControls()
        {
            RearrangeControls(rules: extractRules, addButton: extractRulesAddButton);
        }


        /// <summary>
        /// Update the @c NotesParser using the rules' current state.
        /// </summary>
        
        internal void RegisterChanges()
        {
            if (configLoading)
                return;

            // Need to tell the parser object that the rules have changed.
            parser.UpdateConfig(configObj: config, updateOriginalSourceColumn: false);

            SetRunButtonStatus();
        }

        private void RightRadioButton_Click(object sender, EventArgs e)
        {
            if (configLoading)
                return;

            // Change the left/right preference.
            config.NewColumnLocation = InsertSide.Right;
            parser.UpdateConfig(config);
        }

        /// <summary>
        /// Callback for when the @c runButton is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        /// 
        private void RunButton_Click(object sender, EventArgs e)
        {
            log.Debug("Run button clicked.");

            // Check to see if any enabled extract rules are missing capture groups.
            // (If so, why are we running them?)
            if (CheckExtractRulesUseful())
            {
                ShowCleaningResult();
                ShowDateConversionResult();
                ShowExtractResult();
            }
        }

        /// <summary>
        /// Scrape the GUI & save the rules' current state.
        /// <param name="filename">Name of file to save. Default: configFilename</param>
        /// </summary>

        private void Save(string filename = "")
        {
            // Scrape the GUI & get the latest config.
            parser.UpdateConfig(configObj: config, updateOriginalSourceColumn: false);

            if (String.IsNullOrEmpty(filename))
            {
                filename = configFilename;
            }
                
            using (var writer = new System.IO.StreamWriter(filename))
            {
                var serializer = new XmlSerializer(typeof(NotesConfig));
                serializer.Serialize(writer, config);
                writer.Flush();
            }
        }

        /// <summary>
        /// Scrape the GUI & save the rules' current state to a new file.
        /// </summary>
        
        private void SaveAs()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "XML file|*.xml";
            dialog.Title = "Save config file.";
            dialog.ShowDialog();

            if (dialog.FileName != "")
            {
                Save(filename: dialog.FileName);
            }
        }

        /// <summary>
        /// Callback for when the @c saveAsToolStripMenuItem is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAs();
        }

        /// <summary>
        /// Callback for when the @c saveButton is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void saveButton_Click(object sender, EventArgs e)
        {
            SaveAs();
        }

        /// <summary>
        /// Callback for when the @c saveToolStripMenuItem is clicked.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void SetNewColumnLocation(InsertSide insertSide)
        {
            if (insertSide == InsertSide.Right)
            {
                leftRadioButton.Checked = false;
                rightRadioButton.Checked = true;
            }
            else
            {
                leftRadioButton.Checked = true;
                rightRadioButton.Checked = false;
            }
        }

        /// <summary>
        /// Decide if the @c runButton can be enabled.
        /// </summary>
        
        private void SetRunButtonStatus()
        {
            // If NO cleaning rules, NO date conversion rule and NO extract rules, then the button should be disabled.
            if (
                config.HasCleaningRules()
                || config.HasDateConversionRule()
                || config.HasExtractRules()
            )
            {
                runButton.Enabled = true;
                runButton.BackColor = Color.White;
                runButton.ForeColor = Color.DarkBlue;
                log.Debug("Enabling run button.");
            }
            else
            {
                runButton.Enabled = false;
                runButton.BackColor = Color.Gray;
                runButton.ForeColor = Color.LightGray;
                log.Debug("Disabling run button.");
            }
        }

        /// <summary>
        /// Rolls the @c sourceColumnListBox to the desired column name.
        /// </summary>
        /// <param name="sourceColumn">Name of selected column on which to run these rules</param>
        
        private void SetSourceColumn(string sourceColumn)
        {
            try
            {
                sourceColumnListBox.SelectedItem = sourceColumn;
            }
            catch (Exception) { }
        }

        /// <summary>
        /// Updates the GUI to use the latest cleaning rules.
        /// </summary>
        
        private void ShowCleaningResult()
        {
            log.Debug("Showing cleaning results.");

            // Need to tell the parser object that the rules have changed.
            parser.UpdateConfig(configObj: config, updateOriginalSourceColumn: false);

            if (config.HasCleaningRules())
            {
                parser.Clean();
            }
        }

        /// <summary>
        /// Updates the GUI to use the latest date conversion rules.
        /// </summary>
        
        private void ShowDateConversionResult()
        {
            log.Debug("Showing date conversion results.");

            // Need to tell the parser object that the rules have changed.
            parser.UpdateConfig(configObj: config, updateOriginalSourceColumn: false);

            if (config.HasDateConversionRule())
            {
                parser.ConvertDatesToStandardFormat();
            }
        }

        /// <summary>
        /// Updates the GUI to use the latest data extraction rules.
        /// </summary>
        
        private void ShowExtractResult()
        {
            log.Debug("Showing extraction results.");

            // Need to tell the parser object that the rules have changed.
            parser.UpdateConfig(configObj: config, updateOriginalSourceColumn: false);

            if (config.HasExtractRules())
            {
                parser.Extract();
                parser.SaveRevised();
            }
            else
            {
                // Still need to reset row & remove Status Form.
                parser.ResetAfterProcessing();
            }
        }

        /// <summary>
        /// Updates the GUI to run the rules on just the selected rows.
        /// </summary>
        
        private void ShowSelectedRows()
        {
            ProcessingRowsSelection rowSelection = parser.WhichRowsToProcess();
            ShowSelectedRows(rowSelection);
        }

        /// <summary>
        /// Updates the GUI to run the rules on just the selected rows.
        /// </summary>
        /// <param name="rowSelection">Selected rows</param>
        
        private void ShowSelectedRows(ProcessingRowsSelection rowSelection)
        {
            Excel.Range selectedRows = rowSelection.GetRows();

            string selectionReason = rowSelection.GetReason();

            if (rowSelection.AllRows())
            {
                selectedRowsLabel.Text = "Processing ALL rows.";
            }
            else
            {
                try
                {
                    int minRow = selectedRows[0].Row + 1;
                    int maxRow = selectedRows[selectedRows.Count - 1].Row + 1;

                    if (selectedRows.Count == 1)
                    {
                        selectedRowsLabel.Text =
                            "Processing "
                            + selectedRows.Count.ToString()
                            + " row: "
                            + minRow.ToString();
                    }
                    else if (selectedRows.Count > 1)
                    {
                        selectedRowsLabel.Text =
                            "Processing " + selectedRows.Count.ToString() + " rows:";
                        selectedRowsLabel.Text +=
                            Environment.NewLine
                            + "["
                            + minRow.ToString()
                            + ":"
                            + maxRow.ToString()
                            + "]";
                    }
                }
                catch (System.Runtime.InteropServices.COMException) { }
            }

            if (!string.IsNullOrEmpty(selectionReason))
            {
                selectedRowsLabel.Text += Environment.NewLine + selectionReason;
            }
        }

        /// <summary>
        /// Callback for when the @c sourceColumnListBox selection is changed.
        /// </summary>
        /// <param name="sender">Whatever object trigged this callback.</param>
        /// <param name="e">The EventArgs that accompanied this callback.</param>
        
        private void sourceColumnListBox_Selected(object sender, EventArgs e)
        {
            if (configLoading)
                return;

            // Restore what was in the source column BEFORE we change the column.
            parser.ResetWorksheet();

            // Change the source column...
            string selectedColumnName = sourceColumnListBox.SelectedItem.ToString();
            config.SourceColumnName = selectedColumnName;
            parser.UpdateConfig(config);

            // ...then save its original entries.
            parser.SaveOriginalSourceColumn();

            // Show results of rules on NEW source column.
            Trace.WriteLine(
                "Source column selection changed. Calling ShowCleaningResult() and ShowExtractResult()."
            );
        }
    }
}
