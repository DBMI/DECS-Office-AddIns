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
            this.progressGroupBox = new System.Windows.Forms.GroupBox();
            this.progressBarLabel = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.statusGroupBox = new System.Windows.Forms.GroupBox();
            this.statusLabel = new System.Windows.Forms.Label();
            this.processingStopButton = new System.Windows.Forms.Button();
            this.progressGroupBox.SuspendLayout();
            this.statusGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // progressGroupBox
            // 
            this.progressGroupBox.Controls.Add(this.progressBarLabel);
            this.progressGroupBox.Controls.Add(this.progressBar);
            this.progressGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.progressGroupBox.Location = new System.Drawing.Point(98, 68);
            this.progressGroupBox.Margin = new System.Windows.Forms.Padding(4);
            this.progressGroupBox.Name = "progressGroupBox";
            this.progressGroupBox.Padding = new System.Windows.Forms.Padding(4);
            this.progressGroupBox.Size = new System.Drawing.Size(416, 106);
            this.progressGroupBox.TabIndex = 10;
            this.progressGroupBox.TabStop = false;
            this.progressGroupBox.Text = "Progress";
            // 
            // progressBarLabel
            // 
            this.progressBarLabel.AutoSize = true;
            this.progressBarLabel.Location = new System.Drawing.Point(136, 27);
            this.progressBarLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.progressBarLabel.Name = "progressBarLabel";
            this.progressBarLabel.Size = new System.Drawing.Size(97, 15);
            this.progressBarLabel.TabIndex = 9;
            this.progressBarLabel.Text = "Applying rules";
            this.progressBarLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(28, 63);
            this.progressBar.Margin = new System.Windows.Forms.Padding(4);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(358, 28);
            this.progressBar.TabIndex = 8;
            // 
            // statusGroupBox
            // 
            this.statusGroupBox.Controls.Add(this.statusLabel);
            this.statusGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.statusGroupBox.Location = new System.Drawing.Point(98, 212);
            this.statusGroupBox.Margin = new System.Windows.Forms.Padding(4);
            this.statusGroupBox.Name = "statusGroupBox";
            this.statusGroupBox.Padding = new System.Windows.Forms.Padding(4);
            this.statusGroupBox.Size = new System.Drawing.Size(416, 101);
            this.statusGroupBox.TabIndex = 11;
            this.statusGroupBox.TabStop = false;
            this.statusGroupBox.Text = "Status";
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.statusLabel.Location = new System.Drawing.Point(24, 43);
            this.statusLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(130, 15);
            this.statusLabel.TabIndex = 0;
            this.statusLabel.Text = "status of the apparatus";
            // 
            // processingStopButton
            // 
            this.processingStopButton.BackColor = System.Drawing.Color.White;
            this.processingStopButton.Location = new System.Drawing.Point(278, 335);
            this.processingStopButton.Name = "processingStopButton";
            this.processingStopButton.Size = new System.Drawing.Size(75, 33);
            this.processingStopButton.TabIndex = 12;
            this.processingStopButton.Text = "Stop";
            this.processingStopButton.UseVisualStyleBackColor = false;
            this.processingStopButton.Click += new System.EventHandler(this.processingStopButton_Click);
            // 
            // StatusForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(638, 384);
            this.Controls.Add(this.processingStopButton);
            this.Controls.Add(this.statusGroupBox);
            this.Controls.Add(this.progressGroupBox);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "StatusForm";
            this.Text = "Processing Status";
            this.progressGroupBox.ResumeLayout(false);
            this.progressGroupBox.PerformLayout();
            this.statusGroupBox.ResumeLayout(false);
            this.statusGroupBox.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox progressGroupBox;
        private System.Windows.Forms.Label progressBarLabel;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.GroupBox statusGroupBox;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.Button processingStopButton;
    }
}