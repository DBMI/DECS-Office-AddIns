using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DECS_Excel_Add_Ins
{
    public partial class ChooseTimeThresholdsForm: Form
    {
        public int highUpperThresholdValue;
        public ThresholdCondition highUpperThresholdCondition;
        public int mediumUpperThresholdValue;
        public ThresholdCondition mediumUpperThresholdCondition;
        
        private Dictionary<string, ThresholdCondition> thresholdConditionDict;

        public ChooseTimeThresholdsForm(FollowUpTimeframeThresholds thresholds)
        {
            InitializeComponent();

            // How to look up Enum value from description.
            InitializeThresholdConditionDictionary();

            // Populate & set condition & value listboxes.
            InitializeListBoxes(thresholds);

            // Initialize threshold conditions.
            InitializeThresholdConditions(thresholds);
        }

        public void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        public void HighUrgencyUpperThresholdConditionChanged(object sender, EventArgs e)
        {
            if (thresholdConditionDict.ContainsKey(highUpperThresholdConditionListBox.SelectedItem.ToString()))
            {
                // Update the state variable.
                highUpperThresholdCondition = thresholdConditionDict[highUpperThresholdConditionListBox.SelectedItem.ToString()];

                // Synch the two thresholds.
                mediumLowerThresholdConditionLabel.Text = OtherChoice(highUpperThresholdCondition).GetDescription();
            }
        }

        public void HighUrgencyUpperThresholdValueChanged(object sender, EventArgs e)
        {
            // Don't allow this threshold to be greater than the medium urgency upper threshold.
            highUpperNumericUpDown.Value = Math.Max(1, Math.Min(highUpperNumericUpDown.Value, mediumUpperThresholdValue - 1));

            // Synch the two thresholds.
            mediumLowerThresholdValueLabel.Text = highUpperNumericUpDown.Value.ToString();

            // Update the state variable.
            highUpperThresholdValue = (int)Math.Round(highUpperNumericUpDown.Value);
        }

        private void InitializeThresholdConditionDictionary()
        {
            thresholdConditionDict = new Dictionary<string, ThresholdCondition>();

            foreach (ThresholdCondition condition in Enum.GetValues(typeof(ThresholdCondition)))
            {
                string descr = condition.GetDescription();
                thresholdConditionDict.Add(descr, condition);
            }
        }

        private void InitializeListBoxes(FollowUpTimeframeThresholds thresholds)
        {
            List<string> conditions1 = new List<string>(thresholdConditionDict.Keys);

            // Set the medium urgency values before the high urgency values to avoid a critical race.
            mediumUpperThresholdConditionListBox.DataSource = conditions1;
            mediumUpperNumericUpDown.Value = thresholds.mediumUrgencyUpperThresholdValue;

            List<string> conditions2 = new List<string>(thresholdConditionDict.Keys);

            highUpperThresholdConditionListBox.DataSource = conditions2;
            highUpperNumericUpDown.Value = thresholds.highUrgencyUpperThresholdValue;
        }

        private void InitializeThresholdConditions(FollowUpTimeframeThresholds thresholds)
        {
            routineLowerThresholdConditionLabel.Text = OtherChoice(thresholds.mediumUrgencyUpperThresholdCondition).GetDescription();
            mediumUpperThresholdConditionListBox.SelectedItem = thresholds.mediumUrgencyUpperThresholdCondition.GetDescription();
            mediumLowerThresholdConditionLabel.Text = OtherChoice(thresholds.highUrgencyUpperThresholdCondition).GetDescription();
            highUpperThresholdConditionListBox.SelectedItem = thresholds.highUrgencyUpperThresholdCondition.GetDescription();
        }

        public void MediumUpperThresholdConditionChanged(object sender, EventArgs e)
        {
            if (thresholdConditionDict.ContainsKey(mediumUpperThresholdConditionListBox.SelectedItem.ToString()))
            {
                // Update the state variable.
                mediumUpperThresholdCondition = thresholdConditionDict[mediumUpperThresholdConditionListBox.SelectedItem.ToString()];

                // Synch the two thresholds.
                routineLowerThresholdConditionLabel.Text = OtherChoice(mediumUpperThresholdCondition).GetDescription();
            }
        }

        public void MediumUpperThresholdValueChanged(object sender, EventArgs e)
        {
            // Don't allow this threshold to be below the high urgency upper threshold.
            mediumUpperNumericUpDown.Value = Math.Max(mediumUpperNumericUpDown.Value, highUpperThresholdValue + 1);

            // Synch the two thresholds.
            routineLowerThresholdValueLabel.Text = mediumUpperNumericUpDown.Value.ToString();

            // Update the state variable.
            mediumUpperThresholdValue = (int)Math.Round(mediumUpperNumericUpDown.Value);
        }

        private string OtherChoice(string condition)
        {
            if (condition == "<")
            {
                return "≤";
            }

            return "<";
        }

        private ThresholdCondition OtherChoice(ThresholdCondition condition)
        {
            if (condition == ThresholdCondition.lt)
            {
                return ThresholdCondition.lte;
            }

            return ThresholdCondition.lt;
        }

        public void RunButton_Click(object sender, EventArgs e)
        {
            highUpperThresholdCondition = ThresholdCondition.Unknown;
            mediumUpperThresholdCondition = ThresholdCondition.Unknown;

            highUpperThresholdValue = (int) Math.Round(highUpperNumericUpDown.Value);

            if (thresholdConditionDict.ContainsKey(highUpperThresholdConditionListBox.SelectedItem.ToString()))
            {
                highUpperThresholdCondition = thresholdConditionDict[highUpperThresholdConditionListBox.SelectedItem.ToString()];
            }

            mediumUpperThresholdValue = (int)Math.Round(mediumUpperNumericUpDown.Value);

            if (thresholdConditionDict.ContainsKey(mediumUpperThresholdConditionListBox.SelectedItem.ToString()))
            {
                mediumUpperThresholdCondition = thresholdConditionDict[mediumUpperThresholdConditionListBox.SelectedItem.ToString()];
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
