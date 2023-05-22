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

namespace DECS_Excel_Add_Ins
{
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

        internal DefineRulesForm(NotesParser parser)
        {
            this.parser = parser;
            InitializeComponent();
            PopulateSourceColumnListBox();
            AddCleaningRule();
            AddExtractRule();
            this.parser.AssignWorksheetChangedCallback(ShowSelectedRows);
            ShowSelectedRows();
        }
        private void AddCleaningRule(CleaningRule rule = null, bool updateConfig = true)
        {
            // How many do we have now?
            int nextIndex = NumRulesThisType(parent: cleaningRulesPanel);
            int panelY = PANEL_Y + (Y_STEP * nextIndex);

            CleaningRuleGui cleaningRuleGui = new CleaningRuleGui(
                x: PANEL_X, 
                y: panelY, 
                index: nextIndex, 
                parent: cleaningRulesPanel,
                notesConfig: config,
                updateConfig: updateConfig);

            // Tell the cleaning rule panel to let us know when ITS parent (RuleGui)'s
            // Delete button is pressed.
            cleaningRuleGui.AssignExternalDelete(DeleteCleaningRule);

            // Have the cleaning rule panel tell us when text changes so
            // we can invoke the ShowCleaningResult method.
            cleaningRuleGui.AssignExternalRuleChanged(RegisterChanges);
            cleaningRuleGui.Populate(rule);

            // Keep a list.
            cleaningRules.Add(cleaningRuleGui);

            // Line everything up.
            RearrangeCleaningControls();
        }
        private void AddExtractRule(ExtractRule rule = null, bool updateConfig = true)
        {
            // How many do we have now?
            int nextIndex = NumRulesThisType(parent: extractRulesPanel);
            int panelY = PANEL_Y + (Y_STEP * nextIndex);

            ExtractRuleGui extractRuleGui = new ExtractRuleGui(
                x: PANEL_X,
                y: panelY,
                index: nextIndex,
                parent: extractRulesPanel,
                notesConfig: config,
                updateConfig: updateConfig);

            // Tell the cleaning rule panel to let us know when ITS parent (RuleGui)'s
            // Delete button is pressed.
            extractRuleGui.AssignExternalDelete(DeleteExtractRule);

            // Have the extract rule panel tell us when text changes so
            // we can invoke the ShowExtractResult method.
            extractRuleGui.AssignExternalRuleChanged(RegisterChanges);
            extractRuleGui.Populate(rule);

            // Keep a list.
            extractRules.Add(extractRuleGui);

            // Line everything up.
            RearrangeExtractControls();
        }
        private List<RuleGui> AllRules()
        {
            List<RuleGui> rules = CleaningRules();
            rules.AddRange(ExtractRules());
            return rules;
        }        
        // https://stackoverflow.com/a/12127025/18749636
        private List<RuleGui> CleaningRules()
        {
            List<Panel> cleaningRuleGuis = cleaningRulesPanel.Controls.OfType<Panel>().ToList();

            // Assemble the list of RuleGui objects to which these Panels belong.
            List<RuleGui> rules = cleaningRuleGuis.Select(o => (RuleGui)o.Tag).ToList();
            return rules;
        }
        private void cleaningRulesAddButton_Click(object sender, EventArgs e)
        {
            AddCleaningRule();
        }
        private void clearButton_Click(object sender, EventArgs e)
        {
            DeleteAllRules();

            // Wipe out the accumulated object.
            this.config = new NotesConfig();
            this.parser.UpdateConfig(this.config);

            // Restore any cleaning/extraction done.
            this.parser.ResetWorksheet();
        }
        private void DeleteAllRules()
        {
            List<RuleGui> rules = AllRules();

            // Start with the highest indices & delete in descending order.
            // (So we don't try to delete objects which have already changed their index.)
            rules.OrderByDescending(x => x.Index()).ToList().ForEach(r => r.Delete());
        }
        internal void DeleteCleaningRule(RuleGui cleaningRuleGui)
        {
            // Remove this RuleGui from the controls.
            cleaningRulesPanel.Controls.Remove(cleaningRuleGui.PanelObj);

            // Remove from our list of CleaningRuleGui objects.
            cleaningRules.RemoveAll(r => r == cleaningRuleGui);

            cleaningRuleGui.PanelObj.Dispose();

            RearrangeCleaningControls();
        }
        internal void DeleteExtractRule(RuleGui extractRuleGui)
        {
            // Remove this RuleGui from the controls.
            extractRulesPanel.Controls.Remove(extractRuleGui.PanelObj);

            // Remove from our list of CleaningRuleGui objects.
            extractRules.RemoveAll(r => r == extractRuleGui);

            extractRuleGui.PanelObj.Dispose();

            RearrangeExtractControls();
        }
        private void discardButton_Click(object sender, EventArgs e)
        {
            Close();
        }
        private List<RuleGui> ExtractRules()
        {
            List<Panel> extractRuleGuis = extractRulesPanel.Controls.OfType<Panel>().ToList();

            // Assemble the list of RuleGui objects to which these Panels belong.
            List<RuleGui> rules = extractRuleGuis.Select(o => (RuleGui)o.Tag).ToList();
            return rules;
        }
        private void extractRulesAddButton_Click(object sender, EventArgs e)
        {
            AddExtractRule();
        }
        private List<string> GetAvailableColumnNames()
        {
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            List<string> columnNames = Utilities.GetColumnNames(wksheet);
            columnNames.Sort();
            return columnNames;
        }
        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.configFilename = NotesConfig.ChooseConfigFile();
            NotesConfig configLoaded = NotesConfig.ReadConfigFile(this.configFilename);

            if (configLoaded != null)
            {
                // Temporarily disable all other callbacks until we're done loading.
                configLoading = true;

                // Initialize the NotesConfig object.
                this.config = configLoaded;

                // Link it to the Parser.
                this.parser.UpdateConfig(this.config);

                // Clear what's already here.
                DeleteAllRules();

                // Create & populate boxes, but without changing the config variable,
                // because that's already set.
                foreach(CleaningRule rule in this.config.CleaningRules)
                {
                    AddCleaningRule(rule: rule, updateConfig: false);
                }

                foreach (ExtractRule rule in this.config.ExtractRules)
                {
                    AddExtractRule(rule: rule, updateConfig: false);
                }

                // Roll the source column listbox to the proper column name.
                SetSourceColumn(this.config.SourceColumn);
            }

            // Regard all further alarms.
            configLoading = false;

            Trace.WriteLine("Config file loaded; calling ShowCleaningResult(), ShowExtractResult().");
            ShowCleaningResult();
            ShowExtractResult();
        }
        private void Lock()
        {
            foreach (RuleGui rule in cleaningRules)
            {
                rule.Lock();
            }

            foreach (RuleGui rule in extractRules)
            {
                rule.Lock();
            }
        }
        private int NumPanelsThisType(Panel parent)
        {
            List<Panel> panels = parent.Controls.OfType<Panel>().ToList();
            return panels.Count;
        }
        private int NumRulesThisType(Panel parent)
        {
            List<Panel> panels = parent.Controls.OfType<Panel>().ToList();
            List<RuleGui> rulesThisType = panels.Select(o => (RuleGui)o.Tag).ToList();
            return rulesThisType.Count;
        }
        private void PopulateSourceColumnListBox()
        {
            sourceColumnListBox.DataSource = null;
            sourceColumnListBox.Items.Clear();
            sourceColumnListBox.DataSource = GetAvailableColumnNames();
        }
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
        internal void RearrangeCleaningControls()
        {
            RearrangeControls(rules: this.cleaningRules, addButton: cleaningRulesAddButton);
        }
        internal void RearrangeExtractControls()
        {
            RearrangeControls(rules: this.extractRules, addButton: extractRulesAddButton);
        }
        internal void RegisterChanges()
        {
            if (configLoading) return;

            // Need to tell the parser object that the rules have changed.
            this.parser.UpdateConfig(configObj: this.config, updateOriginalSourceColumn: false);
        }
        private void runButton_Click(object sender, EventArgs e)
        {
            ShowCleaningResult();
            ShowExtractResult();
        }
        private void Save()
        {
            using (var writer = new System.IO.StreamWriter(configFilename))
            {
                var serializer = new XmlSerializer(typeof(NotesConfig));
                serializer.Serialize(writer, this.config);
                writer.Flush();
            }
        }
        private void SaveAs()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "XML file|*.xml";
            dialog.Title = "Save config file.";
            dialog.ShowDialog();

            if (dialog.FileName != "")
            {
                using (var writer = new System.IO.StreamWriter(dialog.FileName))
                {
                    var serializer = new XmlSerializer(typeof(NotesConfig));
                    serializer.Serialize(writer, this.config);
                    writer.Flush();
                }
            }
        }
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAs();
        }
        private void saveButton_Click(object sender, EventArgs e)
        {
            SaveAs();
        }
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Save();
        }
        private void SetSourceColumn(string sourceColumn)
        {
            try
            {
                sourceColumnListBox.SelectedItem = sourceColumn;
            }
            catch (Exception)
            {
            }
        }
        private void ShowCleaningResult()
        {
            // Need to tell the parser object that the rules have changed.
            this.parser.UpdateConfig(configObj: this.config, updateOriginalSourceColumn: false);

            if (this.config.HasCleaningRules())
            {
                this.parser.Clean();
            }            
        }
        private void ShowExtractResult()
        {
            // Need to tell the parser object that the rules have changed.
            this.parser.UpdateConfig(configObj: this.config, updateOriginalSourceColumn: false);

            if (this.config.HasExtractRules())
            {
                this.parser.Extract();
                this.parser.SaveRevised();
            }
            else
            {
                // Still need to reset row & remove Status Form.
                this.parser.ResetAfterProcessing();
            }
        }
        private void ShowSelectedRows()
        {
            ProcessingRowsSelection rowSelection = this.parser.WhichRowsWillBeProcessed();
            ShowSelectedRows(rowSelection);
        }
        private void ShowSelectedRows(ProcessingRowsSelection rowSelection)
        {
            List<int> selectedRows = rowSelection.GetRows();
            string selectionReason = rowSelection.GetReason();

            if (rowSelection.AllRows())
            {
                this.selectedRowsLabel.Text = "Processing ALL rows.";
            }
            else
            {
                int minRow = selectedRows.Min();
                int maxRow = selectedRows.Max();
                
                if (selectedRows.Count == 1)
                {
                    this.selectedRowsLabel.Text = "Processing " + selectedRows.Count.ToString() + " row: " + minRow.ToString();
                }
                else if (selectedRows.Count > 1)
                {
                    this.selectedRowsLabel.Text = "Processing " + selectedRows.Count.ToString() + " rows:";
                    this.selectedRowsLabel.Text += Environment.NewLine + "[" + minRow.ToString() + ":" + maxRow.ToString() + "]";
                }
            }

            if (selectionReason != string.Empty)
            {
                this.selectedRowsLabel.Text += Environment.NewLine + selectionReason;
            }
        }
        private void sourceColumnListBox_Selected(object sender, EventArgs e)
        {
            if (configLoading) return;

            // Restore what was in the source column BEFORE we change the column.
            this.parser.ResetWorksheet();

            // Change the source column...
            string selectedColumn = sourceColumnListBox.SelectedValue.ToString();
            this.config.SourceColumn = selectedColumn;
            this.parser.UpdateConfig(this.config);

            // ...then save its original entries.
            this.parser.SaveOriginalSourceColumn();

            // Show results of rules on NEW source column.
            Trace.WriteLine("Source column selection changed. Calling ShowCleaningResult() and ShowExtractResult().");
            ShowCleaningResult();
            ShowExtractResult();
        }
        private void Unlock()
        {
            foreach (RuleGui rule in cleaningRules)
            {
                rule.Unlock();
            }

            foreach (RuleGui rule in extractRules)
            {
                rule.Unlock();
            }
        }
    }
}