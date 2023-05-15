using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace DECS_Excel_Add_Ins
{
    internal partial class DefineRules : Form
    {
        private NotesConfig config = new NotesConfig();
        private string configFilename = string.Empty;

        private const int BUTTON_HEIGHT = 30;
        private readonly Font BUTTON_FONT = new Font("Microsoft San Serif", 14.25f, FontStyle.Bold);
        private const int BUTTON_WIDTH = 40;
        private const int BUTTON_X = 715;
        private readonly int BUTTON_Y_OFFSET = (int)(RulePanel.Height() - BUTTON_HEIGHT) / 2;

        private const int PANEL_X = 50;
        private const int PANEL_Y = 10;
        private readonly int Y_STEP = RulePanel.Height();

        private NotesParser parser;

        internal DefineRules(NotesParser parser)
        {
            this.parser = parser;
            InitializeComponent();
            PopulateSourceColumnListBox();
            AddCleaningRule();
            AddExtractRule();
        }
        private void AddCleaningRule(CleaningRule rule = null, bool updateConfig = true)
        {
            // How many do we have now?
            List<Panel> cleaningRulePanels = FindPanelsNamed(parent: cleaningRulesPanel, keyword: "cleaningRules");
            int nextIndex = cleaningRulePanels.Count;
            int panelY = PANEL_Y + (Y_STEP * nextIndex);

            CleaningRulePanel cleaningRulePanel = new CleaningRulePanel(
                x: PANEL_X, 
                y: panelY, 
                index: nextIndex, 
                parent: cleaningRulesPanel,
                notesConfig: config,
                updateConfig: updateConfig);

            // Tell the cleaning rule panel to let us know when ITS parent (RulePanel)'s
            // Delete button is pressed.
            cleaningRulePanel.AssignExternalDelete(BumpUpCleaningAddButton);

            // Have the cleaning rule panel tell us when text changes so
            // we can invoke the ShowCleaningResult method.
            cleaningRulePanel.AssignExternalRuleChanged(ShowCleaningResult);
            cleaningRulePanel.Populate(rule);

            // Add bump the Add button to line up below the new panel.
            Point addButtonPosit = new Point(BUTTON_X, panelY + Y_STEP + BUTTON_Y_OFFSET);
            cleaningRulesAddButton.Location = addButtonPosit;
        }
        private void AddExtractRule(ExtractRule rule = null, bool updateConfig = true)
        {
            // How many do we have now?
            List<Panel> extractRulePanels = FindPanelsNamed(parent: extractRulesPanel, keyword: "extractRules");
            int nextIndex = extractRulePanels.Count;
            int panelY = PANEL_Y + (Y_STEP * nextIndex);

            ExtractRulePanel extractRulePanel = new ExtractRulePanel(
                x: PANEL_X,
                y: panelY,
                index: nextIndex,
                parent: extractRulesPanel,
                notesConfig: config,
                updateConfig: updateConfig);

            // Tell the cleaning rule panel to let us know when ITS parent (RulePanel)'s
            // Delete button is pressed.
            extractRulePanel.AssignExternalDelete(BumpUpExtractAddButton);

            // Have the extract rule panel tell us when text changes so
            // we can invoke the ShowExtractResult method.
            extractRulePanel.AssignExternalRuleChanged(ShowExtractResult);
            extractRulePanel.Populate(rule);

            // Add bump the Add button to line up below the new panel.
            Point addButtonPosit = new Point(BUTTON_X, panelY + Y_STEP + BUTTON_Y_OFFSET);
            extractRulesAddButton.Location = addButtonPosit;
        }
        private List<RulePanel> AllRules()
        {
            List<RulePanel> rules = CleaningRules();
            rules.AddRange(ExtractRules());
            return rules;
        }
        internal void BumpUpCleaningAddButton()
        {
            Point addButtonPosit = cleaningRulesAddButton.Location;
            addButtonPosit.Y -= Y_STEP;
            cleaningRulesAddButton.Location = addButtonPosit;
        }
        internal void BumpUpExtractAddButton()
        {
            Point addButtonPosit = extractRulesAddButton.Location;
            addButtonPosit.Y -= Y_STEP;
            extractRulesAddButton.Location = addButtonPosit;
        }
        private List<RulePanel> CleaningRules()
        {
            List<Panel> cleaningRulePanels = cleaningRulesPanel.Controls.OfType<Panel>().ToList();

            // Assemble the list of RulePanel objects to which these Panels belong
            // and invoke their Clear() method.
            List<RulePanel> rules = cleaningRulePanels.Select(o => (RulePanel)o.Tag).ToList();
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
            List<RulePanel> rules = AllRules();

            // Start with the highest indices & delete in descending order.
            // (So we don't try to delete objects who've already changed their index.)
            rules.OrderByDescending(x => x.Index()).ToList().ForEach(r => r.Delete());
        }
        private void discardButton_Click(object sender, EventArgs e)
        {
            Close();
        }
        private List<RulePanel> ExtractRules()
        {
            List<Panel> extractRulePanels = extractRulesPanel.Controls.OfType<Panel>().ToList();

            // Assemble the list of RulePanel objects to which these Panels belong.
            List<RulePanel> rules = extractRulePanels.Select(o => (RulePanel)o.Tag).ToList();
            return rules;
        }
        private void extractRulesAddButton_Click(object sender, EventArgs e)
        {
            AddExtractRule();
        }
        private List<Panel> FindPanelsNamed(Panel parent, string keyword)
        {
            List<Panel> panels = parent.Controls.OfType<Panel>().ToList();
            List<Panel> matchingPanels = panels.Where(b => b.Name.Contains(keyword)).ToList();
            return matchingPanels;
        }        
        private List<string> GetAvailableColumnNames()
        {
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            List<string> columnNames = Utilities.GetColumnNames(wksheet);
            columnNames.Sort();
            return columnNames;
        }
        // We created a NotesParser object without a config object.
        // Now that we're defining a NotesConfig object, let the NotesParser know about it.
        private void InitializeConfig()
        {
            if (config.IsEmpty()) return;

            if (this.parser.HasConfig()) return;

            this.parser.UpdateConfig(this.config);
        }
        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.configFilename = NotesConfig.ChooseConfigFile();
            NotesConfig configLoaded = NotesConfig.ReadConfigFile(this.configFilename);

            if (configLoaded != null)
            {
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

            ShowCleaningResult();
            ShowExtractResult();
        }
        private void PopulateSourceColumnListBox()
        {
            sourceColumnListBox.DataSource = null;
            sourceColumnListBox.Items.Clear();
            sourceColumnListBox.DataSource = GetAvailableColumnNames();
        }
        private void PopulateSourceColumnListBox(string sourceColumn)
        {
            PopulateSourceColumnListBox();
            SetSourceColumn(sourceColumn);
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

            if (this.config.CleaningRules.Count > 0)
            {
                this.parser.Clean();
            }            
        }
        private void ShowExtractResult()
        {
            // Need to tell the parser object that the rules have changed.
            this.parser.UpdateConfig(configObj: this.config, updateOriginalSourceColumn: false);

            if (this.config.ExtractRules.Count > 0)
            {
                this.parser.Extract();
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
        private void sourceColumnListBox_Selected(object sender, EventArgs e)
        {
            // Restore what was in the source column BEFORE we change the column.
            this.parser.ResetWorksheet();

            // Change the source column...
            string selectedColumn = sourceColumnListBox.SelectedValue.ToString();
            this.config.SourceColumn = selectedColumn;
            this.parser.UpdateConfig(this.config);

            // ...then save its original entries.
            this.parser.SaveOriginalSourceColumn();

            // Show results of rules on NEW source column.
            ShowCleaningResult();
            ShowExtractResult();
        }
    }
}
