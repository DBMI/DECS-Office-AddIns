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
    public partial class StatusForm : Form
    {
        private Action externalStopAction;

        public StatusForm(Action parentStopAction)
        {
            InitializeComponent();
            externalStopAction = parentStopAction;
        }
        internal void UpdateProgressBar(int percentage)
        {
            if (progressBar.InvokeRequired)
            {
                Action setProgress = delegate { UpdateProgressBar(percentage); };
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
                Action setLabel = delegate { UpdateProgressBarLabel(text); };
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
                Action setLabel = delegate { UpdateStatusLabel(text); };
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
            // Let calling class know user has requested STOP.
            externalStopAction();
        }
    }
}