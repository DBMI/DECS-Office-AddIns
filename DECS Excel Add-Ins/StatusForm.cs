using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using log4net;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Custom form to show processing status.
     */
    public partial class StatusForm : Form
    {
        private int count;
        private CultureInfo culture = new CultureInfo("en-US");
        private Action externalStopAction;
        private int numRepetitions;
        private Stopwatch stopWatch;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_numRepetitions">int: How many rows will we process?</param>
        /// <param name="parentStopAction">Action: What to do if user presses "Stop" button?</param>
        public StatusForm(int _numRepetitions, Action parentStopAction)
        {
            InitializeComponent();
            externalStopAction = parentStopAction;
            log.Debug("Status form instantiated.");
            count = 0;
            numRepetitions = _numRepetitions;
            stopWatch = new Stopwatch();
            stopWatch.Start();
        }

        /// <summary>
        /// Resets status, resets stopwatch.
        /// </summary>
        /// <param name="_numRepetitions"></param>
        internal void Reset(int _numRepetitions)
        {
            count = 0;
            numRepetitions = _numRepetitions;
            stopWatch.Stop();
            stopWatch.Start();
        }

        /// <summary>
        /// Show we've finished processing one row.
        /// </summary>
        /// <param name="increment">int (default = 1)</param>
        internal void UpdateCount(int increment = 1)
        {
            count += increment;

            int progressPercentage = 100;

            if (numRepetitions > 1)
            {
                progressPercentage = 100 * count / numRepetitions;
            }

            UpdateProgressBar(progressPercentage);
            UpdatePredictedCompletion(stopWatch.GetEta(count, numRepetitions));
        }

        /// <summary>
        /// Update the displayed predicted completion time.
        /// </summary>
        /// <param name="timeRemaining">TimeSpan object</param>
        private void UpdatePredictedCompletion(TimeSpan timeRemaining)
        {
            string predictedCompletion =
                "Completion in " + timeRemaining.ToString(@"hh\:mm\:ss", culture);
            predictedCompletionLabel.Text = predictedCompletion;
        }

        /// <summary>
        /// Update the displayed percent progress.
        /// </summary>
        /// <param name="percentage">int</param>
        private void UpdateProgressBar(int percentage)
        {
            if (progressBar.InvokeRequired)
            {
                Action setProgress = delegate
                {
                    UpdateProgressBar(percentage);
                };
                progressBar.Invoke(setProgress);
            }
            else
            {
                progressBar.Value = percentage;
            }

            Application.DoEvents();
        }

        /// <summary>
        /// Display what processing step we're on.
        /// </summary>
        /// <param name="text"></param>
        internal void UpdateProgressBarLabel(string text)
        {
            if (progressBarLabel.InvokeRequired)
            {
                Action setLabel = delegate
                {
                    UpdateProgressBarLabel(text);
                };
                progressBarLabel.Invoke(setLabel);
            }
            else
            {
                progressBarLabel.Text = text;
            }

            Application.DoEvents();
        }

        /// <summary>
        /// Display what processing step we're on.
        /// </summary>
        /// <param name="text"></param>
        internal void UpdateStatusLabel(string text)
        {
            if (statusLabel.InvokeRequired)
            {
                Action setLabel = delegate
                {
                    UpdateStatusLabel(text);
                };
                statusLabel.Invoke(setLabel);
            }
            else
            {
                statusLabel.Text = text;
            }

            Application.DoEvents();
        }

        /// <summary>
        /// Callback for when user pushes "Stop" button. Passes the action to our calling class.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void processingStopButton_Click(object sender, EventArgs e)
        {
            log.Debug("Stop ordered.");

            // Let calling class know user has requested STOP.
            externalStopAction();
        }
    }
}
