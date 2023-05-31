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
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public StatusForm(int numRepetitions, Action parentStopAction)
        {
            InitializeComponent();
            this.externalStopAction = parentStopAction;
            log.Debug("Status form instantiated.");
            this.count = 0;
            this.numRepetitions = numRepetitions;
            this.stopWatch = new Stopwatch();
            this.stopWatch.Start();
        }
        internal void Reset(int numRepetitions)
        {
            this.count = 0;
            this.numRepetitions = numRepetitions;
            this.stopWatch.Stop();
            this.stopWatch.Start();
        }
        internal void UpdateCount(int increment = 1)
        {
            this.count += increment;

            int progressPercentage = 100;

            if (this.numRepetitions > 1)
            {
                progressPercentage = 100 * this.count / this.numRepetitions;
            }

            UpdateProgressBar(progressPercentage);
            UpdatePredictedCompletion(this.stopWatch.GetEta(this.count, this.numRepetitions));
        }
        private void UpdatePredictedCompletion(TimeSpan timeRemaining)
        {
            string predictedCompletion = "Completion in " + timeRemaining.ToString(@"hh\:mm\:ss", this.culture);
            this.predictedCompletionLabel.Text = predictedCompletion;
        }
        private void UpdateProgressBar(int percentage)
        {
            if (this.progressBar.InvokeRequired)
            {
                Action setProgress = delegate { UpdateProgressBar(percentage); };
                this.progressBar.Invoke(setProgress);
            }
            else
            {
                this.progressBar.Value = percentage;
            }

            Application.DoEvents();
        }
        internal void UpdateProgressBarLabel(string text)
        {
            if (this.progressBarLabel.InvokeRequired)
            {
                Action setLabel = delegate { UpdateProgressBarLabel(text); };
                this.progressBarLabel.Invoke(setLabel);
            }
            else
            {
                this.progressBarLabel.Text = text;
            }

            Application.DoEvents();
        }
        internal void UpdateStatusLabel(string text)
        {
            if (this.statusLabel.InvokeRequired)
            {
                Action setLabel = delegate { UpdateStatusLabel(text); };
                this.statusLabel.Invoke(setLabel);
            }
            else
            {
                this.statusLabel.Text = text;
            }

            Application.DoEvents();
        }
        private void processingStopButton_Click(object sender, EventArgs e)
        {
            log.Debug("Stop ordered.");

            // Let calling class know user has requested STOP.
            this.externalStopAction();
        }
    }
}