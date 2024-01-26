namespace DECS_Excel_Add_Ins
{
    partial class StatusForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(StatusForm));
            progressGroupBox = new System.Windows.Forms.GroupBox();
            progressBarLabel = new System.Windows.Forms.Label();
            progressBar = new System.Windows.Forms.ProgressBar();
            statusGroupBox = new System.Windows.Forms.GroupBox();
            statusLabel = new System.Windows.Forms.Label();
            processingStopButton = new System.Windows.Forms.Button();
            predictedCompletionLabel = new System.Windows.Forms.Label();
            progressGroupBox.SuspendLayout();
            statusGroupBox.SuspendLayout();
            SuspendLayout();
            // 
            // progressGroupBox
            // 
            progressGroupBox.Controls.Add(predictedCompletionLabel);
            progressGroupBox.Controls.Add(progressBarLabel);
            progressGroupBox.Controls.Add(progressBar);
            progressGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            progressGroupBox.Location = new System.Drawing.Point(98, 68);
            progressGroupBox.Margin = new System.Windows.Forms.Padding(4);
            progressGroupBox.Name = "progressGroupBox";
            progressGroupBox.Padding = new System.Windows.Forms.Padding(4);
            progressGroupBox.Size = new System.Drawing.Size(416, 136);
            progressGroupBox.TabIndex = 10;
            progressGroupBox.TabStop = false;
            progressGroupBox.Text = "Progress";
            // 
            // progressBarLabel
            // 
            progressBarLabel.AutoSize = true;
            progressBarLabel.Location = new System.Drawing.Point(136, 27);
            progressBarLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            progressBarLabel.Name = "progressBarLabel";
            progressBarLabel.Size = new System.Drawing.Size(97, 15);
            progressBarLabel.TabIndex = 9;
            progressBarLabel.Text = "Applying rules";
            progressBarLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // progressBar
            // 
            progressBar.Location = new System.Drawing.Point(28, 63);
            progressBar.Margin = new System.Windows.Forms.Padding(4);
            progressBar.Name = "progressBar";
            progressBar.Size = new System.Drawing.Size(358, 28);
            progressBar.TabIndex = 8;
            // 
            // statusGroupBox
            // 
            statusGroupBox.Controls.Add(statusLabel);
            statusGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            statusGroupBox.Location = new System.Drawing.Point(98, 212);
            statusGroupBox.Margin = new System.Windows.Forms.Padding(4);
            statusGroupBox.Name = "statusGroupBox";
            statusGroupBox.Padding = new System.Windows.Forms.Padding(4);
            statusGroupBox.Size = new System.Drawing.Size(416, 101);
            statusGroupBox.TabIndex = 11;
            statusGroupBox.TabStop = false;
            statusGroupBox.Text = "Status";
            // 
            // statusLabel
            // 
            statusLabel.AutoSize = true;
            statusLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            statusLabel.Location = new System.Drawing.Point(24, 43);
            statusLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new System.Drawing.Size(130, 15);
            statusLabel.TabIndex = 0;
            statusLabel.Text = "status of the apparatus";
            // 
            // processingStopButton
            // 
            processingStopButton.BackColor = System.Drawing.Color.White;
            processingStopButton.Location = new System.Drawing.Point(278, 335);
            processingStopButton.Name = "processingStopButton";
            processingStopButton.Size = new System.Drawing.Size(75, 33);
            processingStopButton.TabIndex = 12;
            processingStopButton.Text = "Stop";
            processingStopButton.UseVisualStyleBackColor = false;
            processingStopButton.Click += new System.EventHandler(processingStopButton_Click);
            // 
            // predictedCompletionLabel
            // 
            predictedCompletionLabel.AutoSize = true;
            predictedCompletionLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            predictedCompletionLabel.Location = new System.Drawing.Point(25, 108);
            predictedCompletionLabel.Name = "predictedCompletionLabel";
            predictedCompletionLabel.Size = new System.Drawing.Size(149, 15);
            predictedCompletionLabel.TabIndex = 10;
            predictedCompletionLabel.Text = "predicted completion time";
            // 
            // StatusForm
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            BackColor = System.Drawing.Color.White;
            ClientSize = new System.Drawing.Size(638, 384);
            Controls.Add(processingStopButton);
            Controls.Add(statusGroupBox);
            Controls.Add(progressGroupBox);
            Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            Icon = ((System.Drawing.Icon)(resources.GetObject("$Icon")));
            Margin = new System.Windows.Forms.Padding(4);
            Name = "StatusForm";
            Text = "Processing Status";
            progressGroupBox.ResumeLayout(false);
            progressGroupBox.PerformLayout();
            statusGroupBox.ResumeLayout(false);
            statusGroupBox.PerformLayout();
            ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox progressGroupBox;
        private System.Windows.Forms.Label progressBarLabel;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.GroupBox statusGroupBox;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.Button processingStopButton;
        private System.Windows.Forms.Label predictedCompletionLabel;
    }
}