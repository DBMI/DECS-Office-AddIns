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
    public partial class DefineRules : Form
    {
        private NotesConfig config = new NotesConfig();
        private string configFilename = string.Empty;

        private const int BUTTON_HEIGHT = 30;
        private readonly Font BUTTON_FONT = new Font("Microsoft San Serif", 14.25f, FontStyle.Bold);
        private const int BUTTON_WIDTH = 40;
        private const int BUTTON_X = 715;
        private readonly int BUTTON_Y_OFFSET = (int)(RulePanel.Height() - BUTTON_HEIGHT) / 2;

        private const int PANEL_X = 50;
        private const int PANEL_Y = 46;
        private readonly int Y_STEP = RulePanel.Height();

        public DefineRules()
        {
            InitializeComponent();
            PopulateSourceColumnListBox();
            AddCleaningRule();
            AddExtractRule();
        }
        private void AddCleaningRule(CleaningRule rule = null, bool updateConfig = true)
        {
            // How many do we have now?
            List<Panel> cleaningRulePanels = FindPanelsNamed(parent: cleaningRulesGroupBox, keyword: "cleaningRules");
            int nextIndex = cleaningRulePanels.Count;
            int panelY = PANEL_Y + (Y_STEP * nextIndex);

            CleaningRulePanel cleaningRulePanel = new CleaningRulePanel(
                x: PANEL_X, 
                y: panelY, 
                index: nextIndex, 
                parent: cleaningRulesGroupBox,
                notesConfig: config,
                updateConfig: updateConfig);

            // Tell the cleaning rule panel to let us know when ITS parent (RulePanel)'s
            // Delete button is pressed.
            cleaningRulePanel.AssignExternalDelete(BumpUpCleaningAddButton);
            cleaningRulePanel.Populate(rule);

            // Add bump the Add button to line up below the new panel.
            Point addButtonPosit = new Point(BUTTON_X, panelY + Y_STEP + BUTTON_Y_OFFSET);
            cleaningRulesAddButton.Location = addButtonPosit;
        }
        private void AddExtractRule(ExtractRule rule = null, bool updateConfig = true)
        {
            // How many do we have now?
            List<Panel> extractRulePanels = FindPanelsNamed(parent: extractRulesGroupBox, keyword: "extractRules");
            int nextIndex = extractRulePanels.Count;
            int panelY = PANEL_Y + (Y_STEP * nextIndex);

            ExtractRulePanel extractRulePanel = new ExtractRulePanel(
                x: PANEL_X,
                y: panelY,
                index: nextIndex,
                parent: extractRulesGroupBox,
                notesConfig: config,
                updateConfig: updateConfig);

            // Tell the cleaning rule panel to let us know when ITS parent (RulePanel)'s
            // Delete button is pressed.
            extractRulePanel.AssignExternalDelete(BumpUpExtractAddButton);
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
        public void BumpUpCleaningAddButton()
        {
            Point addButtonPosit = cleaningRulesAddButton.Location;
            addButtonPosit.Y -= Y_STEP;
            cleaningRulesAddButton.Location = addButtonPosit;
        }
        public void BumpUpExtractAddButton()
        {
            Point addButtonPosit = extractRulesAddButton.Location;
            addButtonPosit.Y -= Y_STEP;
            extractRulesAddButton.Location = addButtonPosit;
        }
        private List<RulePanel> CleaningRules()
        {
            List<Panel> cleaningRulePanels = cleaningRulesGroupBox.Controls.OfType<Panel>().ToList();

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
            ClearConfigGui();
            DeleteAllRules();

            // Wipe out the accumulated object.
            config = new NotesConfig();
        }
        private void ClearConfigGui()
        {
            List <RulePanel> rules = AllRules();
            rules.ForEach(r => r.Clear());
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
            List<Panel> extractRulePanels = extractRulesGroupBox.Controls.OfType<Panel>().ToList();

            // Assemble the list of RulePanel objects to which these Panels belong.
            List<RulePanel> rules = extractRulePanels.Select(o => (RulePanel)o.Tag).ToList();
            return rules;
        }
        private void extractRulesAddButton_Click(object sender, EventArgs e)
        {
            AddExtractRule();
        }
        private List<Panel> FindPanelsNamed(GroupBox parent, string keyword)
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
        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            configFilename = NotesConfig.ChooseConfigFile();
            NotesConfig configLoaded = NotesConfig.ReadConfigFile(configFilename);

            if (configLoaded != null)
            {
                // Initialize the NotesConfig object.
                config = configLoaded;

                // Clear what's already here.
                ClearConfigGui();
                DeleteAllRules();

                // Create & populate boxes.
                foreach(CleaningRule rule in config.CleaningRules)
                {
                    AddCleaningRule(rule: rule, updateConfig: false);
                }

                foreach (ExtractRule rule in config.ExtractRules)
                {
                    AddExtractRule(rule: rule, updateConfig: false);
                }
            }
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

            try
            {
                sourceColumnListBox.SelectedItem = config.SourceColumn;
            }
            catch (Exception)
            {
            }
        }
        private void Save()
        {
            using (var writer = new System.IO.StreamWriter(configFilename))
            {
                var serializer = new XmlSerializer(typeof(NotesConfig));
                serializer.Serialize(writer, config);
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
                    serializer.Serialize(writer, config);
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
        private void sourceColumnListBox_Selected(object sender, EventArgs e)
        {
            string selectedColumn = sender.ToString();
            config.SourceColumn = selectedColumn;
        }
    }
}
