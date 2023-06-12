using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using Action = System.Action;
using Button = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;
using Font = System.Drawing.Font;
using Panel = System.Windows.Forms.Panel;
using Point = System.Drawing.Point;
using TextBox = System.Windows.Forms.TextBox;
using ToolTip = System.Windows.Forms.ToolTip;
using log4net;
using System.Text.RegularExpressions;
using System.Web.UI.WebControls;

namespace DECS_Excel_Add_Ins
{
    // A set of controls: two textboxes, a delete button and the panel that contains them all.
    internal abstract class RuleGui
    {
        private const int BOX_HEIGHT = 22;
        private readonly Font BOX_FONT = new Font("Microsoft San Serif", 9.75f, FontStyle.Regular);
        private const int LEFT_BOX_WIDTH = 150;
        private const int CENTER_BOX_WIDTH = 840;
        private const int RIGHT_BOX_WIDTH = 195;

        private const int BUTTON_HEIGHT = 30;
        private readonly Font BUTTON_FONT = new Font("Microsoft San Serif", 14.25f, FontStyle.Bold);
        private const int BUTTON_WIDTH = 40;
        private const int BUTTON_X = 1270;
        private readonly int BUTTON_Y_OFFSET = (int)(BOX_HEIGHT - BUTTON_HEIGHT) / 2;

        private const int LEFT_BOX_X = 25;
        private const int CENTER_BOX_X = 195;
        private const int RIGHT_BOX_X = 1055;
        private readonly int BOX_Y = (int)BOX_HEIGHT / 2;

        private readonly Font CHECKBOX_FONT = new Font(
            "Microsoft San Serif",
            7f,
            FontStyle.Regular
        );
        private const int CHECKBOX_X = 2;
        private const int CHECKBOX_WIDTH = 13;

        protected Panel panel;
        private Panel parent;
        protected CheckBox checkBox;
        private Button deleteButton;
        protected TextBox leftTextBox;
        protected TextBox centerTextBox;
        protected TextBox rightTextBox;

        protected int index;
        private string keyword;

        private static int width = 1316;
        private static int height = 44;

        private Action<RuleGui> inheritedClassDeleteAction;
        private Action<RuleGui> parentClassDisableAction;
        private Action<RuleGui> parentClassEnableAction;

        private ToolTip toolTip;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        public RuleGui(int x, int y, int index, Panel parentObj, string ruleType)
        {
            this.index = index; // zero-based
            this.parent = parentObj;
            this.keyword = ruleType;

            this.panel = new Panel();
            this.panel.Height = height;
            this.panel.Location = new Point(x, y);
            this.panel.Name = this.keyword;
            this.panel.Parent = parent;
            this.panel.Tag = this; // So that, if we find the Panel object, we can find its associated RuleGui object.
            this.panel.Width = width;
            this.parent.Controls.Add(this.panel);

            // Create enable checkbox.
            this.checkBox = new CheckBox();
            this.checkBox.Checked = true;
            this.checkBox.Click += CheckBoxClicked;
            this.checkBox.Font = CHECKBOX_FONT;
            this.checkBox.Height = BUTTON_HEIGHT;
            System.Drawing.Size checkBoxSize = this.checkBox.Size;
            checkBoxSize.Width = CHECKBOX_WIDTH;
            this.checkBox.Size = checkBoxSize;
            int checkBoxYoffset = (int)(BOX_HEIGHT - checkBoxSize.Height) / 2;
            Point checkBoxPosit = new Point(CHECKBOX_X, BOX_Y + checkBoxYoffset);
            this.checkBox.Location = checkBoxPosit;
            this.checkBox.Parent = this.panel;
            this.checkBox.Text = "";

            // Create and position boxes.
            this.leftTextBox = new TextBox();
            this.leftTextBox.Parent = this.panel;
            this.leftTextBox.Font = BOX_FONT;
            this.leftTextBox.Height = BOX_HEIGHT;
            Point leftPosit = new Point(LEFT_BOX_X, BOX_Y);
            this.leftTextBox.Location = leftPosit;
            this.leftTextBox.Name = this.keyword + "LeftTextBox";
            this.leftTextBox.Width = LEFT_BOX_WIDTH;

            this.centerTextBox = new TextBox();
            this.centerTextBox.Parent = this.panel;
            this.centerTextBox.Font = BOX_FONT;
            this.centerTextBox.Height = BOX_HEIGHT;
            Point centerPosit = new Point(CENTER_BOX_X, BOX_Y);
            this.centerTextBox.Location = centerPosit;
            this.centerTextBox.Name = this.keyword + "CenterTextBox";
            this.centerTextBox.Width = CENTER_BOX_WIDTH;

            this.rightTextBox = new TextBox();
            this.rightTextBox.Parent = this.panel;
            this.rightTextBox.Font = BOX_FONT;
            this.rightTextBox.Height = BOX_HEIGHT;
            Point rightHandPosit = new Point(RIGHT_BOX_X, BOX_Y);
            this.rightTextBox.Location = rightHandPosit;
            this.rightTextBox.Name = this.keyword + "RightTextBox";
            this.rightTextBox.Width = RIGHT_BOX_WIDTH;

            // Create new delete button.
            this.deleteButton = new Button();
            this.deleteButton.Click += Delete;
            this.deleteButton.Font = BUTTON_FONT;
            this.deleteButton.Height = BUTTON_HEIGHT;
            Point deleteButtonPosit = new Point(BUTTON_X, BOX_Y + BUTTON_Y_OFFSET);
            this.deleteButton.Location = deleteButtonPosit;
            this.deleteButton.Parent = this.panel;
            this.deleteButton.Text = "−";
            this.deleteButton.Width = BUTTON_WIDTH;

            // Add to controls.
            this.panel.Controls.Add(this.leftTextBox);
            this.panel.Controls.Add(this.centerTextBox);
            this.panel.Controls.Add(this.rightTextBox);
            this.panel.Controls.Add(this.deleteButton);
        }

        // This class creates the Delete button and handles disposing of the GUI elements
        // but knows nothing of the NotesConfig object being built.
        // So classes which inherit RuleGui (CleaningRuleGui & ExtractRuleGui)
        // and DO know about NotesConfig need to be able to assign actions that fire
        // when our our Delete button is pressed.
        // Similarly, the DefineRule Class owns the AddButton and needs to move the
        // button up when a rule is deleted, so it provides ITS callback to the
        // CleaningRuleGui and ExtractRuleGui classes to invoke.
        protected void AssignDelete(Action<RuleGui> deleteAction)
        {
            inheritedClassDeleteAction = deleteAction;
        }

        internal void AssignDisable(Action<RuleGui> disableAction)
        {
            parentClassDisableAction = disableAction;
        }

        internal void AssignEnable(Action<RuleGui> enableAction)
        {
            parentClassEnableAction = enableAction;
        }

        private void CheckBoxClicked(object sender, EventArgs e)
        {
            CheckBox checkBox = sender as CheckBox;

            if (checkBox.Checked)
            {
                log.Debug("Checkbox " + this.index.ToString() + " checked.");

                if (parentClassEnableAction != null)
                {
                    parentClassEnableAction(this);
                }
            }
            else
            {
                log.Debug("Checkbox " + this.index.ToString() + " unchecked.");

                if (parentClassDisableAction != null)
                {
                    parentClassDisableAction(this);
                }
            }
        }

        // Implemented in the derived classes because the TextChanged methods are defined there
        // and we temporarily need to disable the TextChanged methods while clearing the textboxes.
        public abstract void Clear();

        public void Delete()
        {
            // Pass the order down the chain to the next panel (until there isn't one).
            RuleGui nextObject = NextRuleGui();

            if (nextObject != null)
            {
                nextObject.MoveUpInLine();
            }

            // Invoke the inherited class' Delete() function, which removes
            // the cleaning rule or extract rule for this index.
            inheritedClassDeleteAction(this);
        }

        private void Delete(object sender, EventArgs e)
        {
            Delete();
        }

        private RuleGui FindNth(int desiredIndex)
        {
            // Find the underlying Panel objects of this rule type.
            List<Panel> panels = parent.Controls.OfType<Panel>().ToList();

            // Assemble the list of RuleGui objects to which these Panels belong.
            List<RuleGui> rules = panels.Select(o => (RuleGui)o.Tag).ToList();

            // Which one has the index we want?
            List<RuleGui> matchingPanels = rules.Where(b => b.Index() == desiredIndex).ToList();

            if (matchingPanels.Count > 0)
            {
                return (RuleGui)matchingPanels[0];
            }

            return null;
        }

        // So calling class can ask how big a RuleGui object is prior to object instantiation.
        public static int Height()
        {
            return height;
        }

        internal int Index()
        {
            return this.index;
        }

        private void MoveUpInLine()
        {
            // Pass the word.
            RuleGui nextInLine = NextRuleGui();

            nextInLine?.MoveUpInLine();

            // Decrement my index.
            this.index -= 1;
        }

        private RuleGui NextRuleGui()
        {
            return FindNth(this.index + 1);
        }

        public Panel PanelObj
        {
            get { return this.panel; }
        }

        internal void ResetLocation(int x, int y)
        {
            this.panel.Location = new Point(x, y);
        }
    }
}
