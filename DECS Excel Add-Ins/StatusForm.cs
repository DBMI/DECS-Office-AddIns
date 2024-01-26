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

        internal void Reset(int _numRepetitions)
        {
            count = 0;
            numRepetitions = _numRepetitions;
            stopWatch.Stop();
            stopWatch.Start();
        }

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

        private void UpdatePredictedCompletion(TimeSpan timeRemaining)
        {
            string predictedCompletion =
                "Completion in " + timeRemaining.ToString(@"hh\:mm\:ss", culture);
            predictedCompletionLabel.Text = predictedCompletion;
        }

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

        private void processingStopButton_Click(object sender, EventArgs e)
        {
            log.Debug("Stop ordered.");

            // Let calling class know user has requested STOP.
            externalStopAction();
        }
    }
}
