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
using Button = System.Windows.Forms.Button;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using TextBox = System.Windows.Forms.TextBox;

namespace DECS_Excel_Add_Ins
{
    // A set of controls: two textboxes, a delete button and the panel that contains them all.
    internal abstract class RuleGui
    {
        private const int BOX_HEIGHT = 22;
        private readonly Font BOX_FONT = new Font("Microsoft San Serif", 9.75f, FontStyle.Regular);
        private const int LEFT_BOX_WIDTH = 1000;
        private const int RIGHT_BOX_WIDTH = 200;

        private const int BUTTON_HEIGHT = 30;
        private readonly Font BUTTON_FONT = new Font("Microsoft San Serif", 14.25f, FontStyle.Bold);
        private const int BUTTON_WIDTH = 40;
        private const int BUTTON_X = 1270;
        private readonly int BUTTON_Y_OFFSET = (int)(BOX_HEIGHT - BUTTON_HEIGHT) / 2;

        private const int leftHandX = 5;
        private const int rightHandX = 1040;
        private readonly int boxY = (int) BOX_HEIGHT/2;

        protected Panel panel;
        private Panel parent;
        protected TextBox leftHandTextBox;
        protected TextBox rightHandTextBox;
        private Button deleteButton;

        protected int index;
        private string keyword;

        private static int width = 1316;
        private static int height = 44;

        private System.Action<RuleGui> inheritedClassDeleteAction;

        public RuleGui(int x, int y, int index, Panel parentObj, string ruleType)
        {
            this.index = index;
            this.parent = parentObj;
            this.keyword = ruleType;

            this.panel = new Panel();
            this.panel.Height = height;
            this.panel.Location = new Point(x, y);
            this.panel.Name = this.keyword;
            this.panel.Parent = parent;
            this.panel.Tag = this;      // So that, if we find the Panel object, we can find its associated RuleGui object.
            this.panel.Width = width;
            this.parent.Controls.Add(this.panel);

            // Create and position boxes.
            this.leftHandTextBox = new System.Windows.Forms.TextBox();
            this.leftHandTextBox.Parent = this.panel;
            this.leftHandTextBox.Font = BOX_FONT;
            this.leftHandTextBox.Height = BOX_HEIGHT;
            Point leftHandPosit = new Point(leftHandX, boxY);
            this.leftHandTextBox.Location = leftHandPosit;
            this.leftHandTextBox.Name = this.keyword + "LeftTextBox";
            this.leftHandTextBox.Width = LEFT_BOX_WIDTH;

            this.rightHandTextBox = new TextBox();
            this.rightHandTextBox.Parent = this.panel;
            this.rightHandTextBox.Font = BOX_FONT;
            this.rightHandTextBox.Height = BOX_HEIGHT;
            Point rightHandPosit = new Point(rightHandX, boxY);
            this.rightHandTextBox.Location = rightHandPosit;
            this.rightHandTextBox.Name = this.keyword + "RightTextBox";
            this.rightHandTextBox.Width = RIGHT_BOX_WIDTH;

            // Create new delete button.
            this.deleteButton = new Button();
            this.deleteButton.Parent = this.panel;
            this.deleteButton.Font = BUTTON_FONT;
            this.deleteButton.Height = BUTTON_HEIGHT;
            Point deleteButtonPosit = new Point(BUTTON_X, boxY + BUTTON_Y_OFFSET);
            this.deleteButton.Location = deleteButtonPosit;
            this.deleteButton.Text = "−";
            this.deleteButton.Width = BUTTON_WIDTH;
            this.deleteButton.Click += Delete;

            // Add to controls.
            this.panel.Controls.Add(this.leftHandTextBox);
            this.panel.Controls.Add(this.rightHandTextBox);
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
        protected void AssignDelete(System.Action<RuleGui> deleteAction)
        {
            inheritedClassDeleteAction = deleteAction;
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
        private void Enable(Control control, bool locked)
        {
            if (control.InvokeRequired)
            {
                System.Action setProgress = delegate { Enable(control, locked); };
                control.Invoke(setProgress);
            }
            else
            {
                control.Enabled = locked;
            }
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
        internal void Lock()
        {
            Enable(this.leftHandTextBox, false);
            Enable(this.rightHandTextBox, false);
            Enable(this.deleteButton, false);
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
        public Panel PanelObj { get { return this.panel; } }

        internal void ResetLocation(int x, int y)
        {
            this.panel.Location = new Point(x, y);
        }
        internal void Unlock()
        {
            Enable(this.leftHandTextBox, true);
            Enable(this.rightHandTextBox, true);
            Enable(this.deleteButton, true);
        }
        public static int Width()
        {
            return width;
        }
    }
}
