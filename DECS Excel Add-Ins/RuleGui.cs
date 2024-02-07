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
    /**
     * @brief A set of controls: two textboxes, a delete button and the panel that contains them all.
     * Meant to be inherited & instantiated by @c CleaningRule and @c ExtractRule classes.
     */
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

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="x">X position of parent @c Panel object</param>
        /// <param name="y">Y position of parent @c Panel object</param>
        /// <param name="_index">zero-based index of this object in list of rules</param>
        /// <param name="parentObj">Parent @Panel object</param>
        /// <param name="ruleType">Is it a cleaning or an extraction rule?</param>
        public RuleGui(int x, int y, int _index, Panel parentObj, string ruleType)
        {
            index = _index; // zero-based
            parent = parentObj;
            keyword = ruleType;

            panel = new Panel();
            panel.Height = height;
            panel.Location = new Point(x, y);
            panel.Name = keyword;
            panel.Parent = parent;
            panel.Tag = this; // So that, if we find the Panel object, we can find its associated RuleGui object.
            panel.Width = width;
            parent.Controls.Add(panel);

            // Create enable checkbox.
            checkBox = new CheckBox();
            checkBox.Checked = true;
            checkBox.Click += CheckBoxClicked;
            checkBox.Font = CHECKBOX_FONT;
            checkBox.Height = BUTTON_HEIGHT;
            System.Drawing.Size checkBoxSize = checkBox.Size;
            checkBoxSize.Width = CHECKBOX_WIDTH;
            checkBox.Size = checkBoxSize;
            int checkBoxYoffset = (int)(BOX_HEIGHT - checkBoxSize.Height) / 2;
            Point checkBoxPosit = new Point(CHECKBOX_X, BOX_Y + checkBoxYoffset);
            checkBox.Location = checkBoxPosit;
            checkBox.Parent = panel;
            checkBox.Text = "";

            // Create and position boxes.
            leftTextBox = new TextBox();
            leftTextBox.Parent = panel;
            leftTextBox.Font = BOX_FONT;
            leftTextBox.Height = BOX_HEIGHT;
            Point leftPosit = new Point(LEFT_BOX_X, BOX_Y);
            leftTextBox.Location = leftPosit;
            leftTextBox.Name = keyword + "LeftTextBox";
            leftTextBox.Width = LEFT_BOX_WIDTH;

            centerTextBox = new TextBox();
            centerTextBox.Parent = panel;
            centerTextBox.Font = BOX_FONT;
            centerTextBox.Height = BOX_HEIGHT;
            Point centerPosit = new Point(CENTER_BOX_X, BOX_Y);
            centerTextBox.Location = centerPosit;
            centerTextBox.Name = keyword + "CenterTextBox";
            centerTextBox.Width = CENTER_BOX_WIDTH;

            rightTextBox = new TextBox();
            rightTextBox.Parent = panel;
            rightTextBox.Font = BOX_FONT;
            rightTextBox.Height = BOX_HEIGHT;
            Point rightHandPosit = new Point(RIGHT_BOX_X, BOX_Y);
            rightTextBox.Location = rightHandPosit;
            rightTextBox.Name = keyword + "RightTextBox";
            rightTextBox.Width = RIGHT_BOX_WIDTH;

            // Create new delete button.
            deleteButton = new Button();
            deleteButton.Click += Delete;
            deleteButton.Font = BUTTON_FONT;
            deleteButton.Height = BUTTON_HEIGHT;
            Point deleteButtonPosit = new Point(BUTTON_X, BOX_Y + BUTTON_Y_OFFSET);
            deleteButton.Location = deleteButtonPosit;
            deleteButton.Parent = panel;
            deleteButton.Text = "−";
            deleteButton.Width = BUTTON_WIDTH;

            // Add to controls.
            panel.Controls.Add(leftTextBox);
            panel.Controls.Add(centerTextBox);
            panel.Controls.Add(rightTextBox);
            panel.Controls.Add(deleteButton);
        }

        /// <summary>
        /// This class creates the Delete button and handles disposing of the GUI elements
        /// but knows nothing of the @c NotesConfig object being built.
        /// So classes which inherit @c RuleGui (@c CleaningRuleGui & @c ExtractRuleGui)
        /// and DO know about @c NotesConfig need to be able to assign actions that fire
        /// when our our Delete button is pressed.
        /// Similarly, the @c DefineRule Class owns the AddButton and needs to move the
        /// button up when a rule is deleted, so it provides ITS callback to the
        /// @c CleaningRuleGui and @c ExtractRuleGui classes to invoke.
        /// </summary>
        /// <param name="deleteAction">@c Action to perform when delete button pressed</param>
        /// 
        protected void AssignDelete(Action<RuleGui> deleteAction)
        {
            inheritedClassDeleteAction = deleteAction;
        }

        /// <summary>
        /// Lets inherited classes @c CleaningRuleGui & @c ExtractRuleGui assign the callback for disable button.
        /// </summary>
        /// <param name="disableAction">Action</param>
        internal void AssignDisable(Action<RuleGui> disableAction)
        {
            parentClassDisableAction = disableAction;
        }

        /// <summary>
        /// Lets inherited classes @c CleaningRuleGui & @c ExtractRuleGui assign the callback for enable button.
        /// </summary>
        /// <param name="enableAction">Action</param>
        internal void AssignEnable(Action<RuleGui> enableAction)
        {
            parentClassEnableAction = enableAction;
        }



        /// <summary>
        /// Callback for when checkbox is clicked.
        /// </summary>
        /// <param name="sender">Object sending us the callback</param>
        /// <param name="e">Callback arguments</param>
        private void CheckBoxClicked(object sender, EventArgs e)
        {
            CheckBox checkBox = sender as CheckBox;

            if (checkBox.Checked)
            {
                log.Debug("Checkbox " + index.ToString() + " checked.");

                if (parentClassEnableAction != null)
                {
                    parentClassEnableAction(this);
                }
            }
            else
            {
                log.Debug("Checkbox " + index.ToString() + " unchecked.");

                if (parentClassDisableAction != null)
                {
                    parentClassDisableAction(this);
                }
            }
        }

        /// <summary>
        /// Implemented in the derived classes because the TextChanged methods are defined there
        /// and we temporarily need to disable the TextChanged methods while clearing the textboxes.
        /// </summary>
        public abstract void Clear();

        /// <summary>
        /// Pass the delete order down the chain to the next panel (until there isn't one).
        /// Then--at the last panel, invoke the inherited class' Delete() function, which removes 
        /// the cleaning rule or extract rule for this index.
        /// </summary>
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

        /// <summary>
        /// Delete callback. Passes it along to the @c Delete() method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Delete(object sender, EventArgs e)
        {
            Delete();
        }

        /// <summary>
        /// Finds a desired @c RuleGui object by index without top-level object maintaining a list.
        /// </summary>
        /// <param name="desiredIndex">Index of @c Panel we want to find</param>
        /// <returns>@c RuleGui</returns>
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

        /// <summary>
        /// So calling class can ask how big a RuleGui object is prior to object instantiation. 
        /// </summary>
        /// <returns>int</returns>
        public static int Height()
        {
            return height;
        }

        /// <summary>
        /// So calling class can for this object's index.
        /// </summary>
        /// <returns>int</returns>
        internal int Index()
        {
            return index;
        }

        /// <summary>
        /// Moves @c Panels up following a rule deletion.
        /// </summary>
        private void MoveUpInLine()
        {
            // Pass the word.
            RuleGui nextInLine = NextRuleGui();

            nextInLine?.MoveUpInLine();

            // Decrement my index.
            index -= 1;
        }

        /// <summary>
        /// Returns the next object.
        /// </summary>
        /// <returns>@c RuleGui</returns>
        private RuleGui NextRuleGui()
        {
            return FindNth(index + 1);
        }

        /// <summary>
        /// Allows for external reference to object's parent @c Panel object.
        /// </summary>
        public Panel PanelObj
        {
            get { return panel; }
        }

        /// <summary>
        /// Allows @c DefineRulesForm to direct this panel to move.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        internal void ResetLocation(int x, int y)
        {
            panel.Location = new Point(x, y);
        }
    }
}
