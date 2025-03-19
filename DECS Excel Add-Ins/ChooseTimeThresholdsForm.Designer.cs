namespace DECS_Excel_Add_Ins
{
    partial class ChooseTimeThresholdsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChooseTimeThresholdsForm));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.highUpperNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.mediumUpperNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.mediumLowerThresholdValueLabel = new System.Windows.Forms.Label();
            this.highUpperThresholdConditionListBox = new System.Windows.Forms.ListBox();
            this.label11 = new System.Windows.Forms.Label();
            this.mediumLowerThresholdConditionLabel = new System.Windows.Forms.Label();
            this.mediumUpperThresholdConditionListBox = new System.Windows.Forms.ListBox();
            this.routineLowerThresholdConditionLabel = new System.Windows.Forms.Label();
            this.routineLowerThresholdValueLabel = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.highUpperNumericUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mediumUpperNumericUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(85, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(310, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Choose Urgency Thresholds";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(74, 89);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 20);
            this.label2.TabIndex = 1;
            this.label2.Text = "High";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(49, 130);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 20);
            this.label3.TabIndex = 2;
            this.label3.Text = "Medium";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(48, 171);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 20);
            this.label4.TabIndex = 3;
            this.label4.Text = "Routine";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(178, 89);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(18, 20);
            this.label5.TabIndex = 4;
            this.label5.Text = "0";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(229, 89);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(29, 20);
            this.label6.TabIndex = 5;
            this.label6.Text = "ΔT";
            // 
            // highUpperNumericUpDown
            // 
            this.highUpperNumericUpDown.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.highUpperNumericUpDown.Location = new System.Drawing.Point(313, 86);
            this.highUpperNumericUpDown.Maximum = new decimal(new int[] {
            52,
            0,
            0,
            0});
            this.highUpperNumericUpDown.Name = "highUpperNumericUpDown";
            this.highUpperNumericUpDown.Size = new System.Drawing.Size(43, 26);
            this.highUpperNumericUpDown.TabIndex = 6;
            this.highUpperNumericUpDown.ValueChanged += new System.EventHandler(this.HighUrgencyUpperThresholdValueChanged);
            // 
            // mediumUpperNumericUpDown
            // 
            this.mediumUpperNumericUpDown.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mediumUpperNumericUpDown.Location = new System.Drawing.Point(313, 128);
            this.mediumUpperNumericUpDown.Maximum = new decimal(new int[] {
            52,
            0,
            0,
            0});
            this.mediumUpperNumericUpDown.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.mediumUpperNumericUpDown.Name = "mediumUpperNumericUpDown";
            this.mediumUpperNumericUpDown.Size = new System.Drawing.Size(43, 26);
            this.mediumUpperNumericUpDown.TabIndex = 9;
            this.mediumUpperNumericUpDown.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.mediumUpperNumericUpDown.ValueChanged += new System.EventHandler(this.MediumUpperThresholdValueChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(229, 130);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(29, 20);
            this.label7.TabIndex = 10;
            this.label7.Text = "ΔT";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(229, 170);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(29, 20);
            this.label8.TabIndex = 11;
            this.label8.Text = "ΔT";
            // 
            // okButton
            // 
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.Location = new System.Drawing.Point(86, 251);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(92, 33);
            this.okButton.TabIndex = 12;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.RunButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelButton.Location = new System.Drawing.Point(270, 251);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(92, 33);
            this.cancelButton.TabIndex = 13;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(378, 89);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(54, 20);
            this.label9.TabIndex = 14;
            this.label9.Text = "weeks";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(378, 130);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(54, 20);
            this.label10.TabIndex = 15;
            this.label10.Text = "weeks";
            // 
            // mediumLowerThresholdValueLabel
            // 
            this.mediumLowerThresholdValueLabel.AutoSize = true;
            this.mediumLowerThresholdValueLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mediumLowerThresholdValueLabel.Location = new System.Drawing.Point(178, 130);
            this.mediumLowerThresholdValueLabel.Name = "mediumLowerThresholdValueLabel";
            this.mediumLowerThresholdValueLabel.Size = new System.Drawing.Size(18, 20);
            this.mediumLowerThresholdValueLabel.TabIndex = 16;
            this.mediumLowerThresholdValueLabel.Text = "0";
            // 
            // highUpperThresholdConditionListBox
            // 
            this.highUpperThresholdConditionListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.highUpperThresholdConditionListBox.FormattingEnabled = true;
            this.highUpperThresholdConditionListBox.ItemHeight = 20;
            this.highUpperThresholdConditionListBox.Location = new System.Drawing.Point(262, 86);
            this.highUpperThresholdConditionListBox.Name = "highUpperThresholdConditionListBox";
            this.highUpperThresholdConditionListBox.Size = new System.Drawing.Size(39, 24);
            this.highUpperThresholdConditionListBox.Sorted = true;
            this.highUpperThresholdConditionListBox.TabIndex = 18;
            this.highUpperThresholdConditionListBox.SelectedValueChanged += new System.EventHandler(this.HighUrgencyUpperThresholdConditionChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(204, 134);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(0, 13);
            this.label11.TabIndex = 19;
            // 
            // mediumLowerThresholdConditionLabel
            // 
            this.mediumLowerThresholdConditionLabel.AutoSize = true;
            this.mediumLowerThresholdConditionLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mediumLowerThresholdConditionLabel.Location = new System.Drawing.Point(203, 130);
            this.mediumLowerThresholdConditionLabel.Name = "mediumLowerThresholdConditionLabel";
            this.mediumLowerThresholdConditionLabel.Size = new System.Drawing.Size(18, 20);
            this.mediumLowerThresholdConditionLabel.TabIndex = 20;
            this.mediumLowerThresholdConditionLabel.Text = "<";
            // 
            // mediumUpperThresholdConditionListBox
            // 
            this.mediumUpperThresholdConditionListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mediumUpperThresholdConditionListBox.FormattingEnabled = true;
            this.mediumUpperThresholdConditionListBox.ItemHeight = 20;
            this.mediumUpperThresholdConditionListBox.Location = new System.Drawing.Point(262, 128);
            this.mediumUpperThresholdConditionListBox.Name = "mediumUpperThresholdConditionListBox";
            this.mediumUpperThresholdConditionListBox.Size = new System.Drawing.Size(39, 24);
            this.mediumUpperThresholdConditionListBox.Sorted = true;
            this.mediumUpperThresholdConditionListBox.TabIndex = 21;
            this.mediumUpperThresholdConditionListBox.SelectedValueChanged += new System.EventHandler(this.MediumUpperThresholdConditionChanged);
            // 
            // routineLowerThresholdConditionLabel
            // 
            this.routineLowerThresholdConditionLabel.AutoSize = true;
            this.routineLowerThresholdConditionLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.routineLowerThresholdConditionLabel.Location = new System.Drawing.Point(205, 170);
            this.routineLowerThresholdConditionLabel.Name = "routineLowerThresholdConditionLabel";
            this.routineLowerThresholdConditionLabel.Size = new System.Drawing.Size(18, 20);
            this.routineLowerThresholdConditionLabel.TabIndex = 22;
            this.routineLowerThresholdConditionLabel.Text = "<";
            // 
            // routineLowerThresholdValueLabel
            // 
            this.routineLowerThresholdValueLabel.AutoSize = true;
            this.routineLowerThresholdValueLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.routineLowerThresholdValueLabel.Location = new System.Drawing.Point(178, 170);
            this.routineLowerThresholdValueLabel.Name = "routineLowerThresholdValueLabel";
            this.routineLowerThresholdValueLabel.Size = new System.Drawing.Size(18, 20);
            this.routineLowerThresholdValueLabel.TabIndex = 23;
            this.routineLowerThresholdValueLabel.Text = "0";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(204, 89);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(18, 20);
            this.label12.TabIndex = 24;
            this.label12.Text = "≤";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(264, 170);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(54, 20);
            this.label13.TabIndex = 25;
            this.label13.Text = "weeks";
            // 
            // ChooseTimeThresholdsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(481, 339);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.routineLowerThresholdValueLabel);
            this.Controls.Add(this.routineLowerThresholdConditionLabel);
            this.Controls.Add(this.mediumUpperThresholdConditionListBox);
            this.Controls.Add(this.mediumLowerThresholdConditionLabel);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.highUpperThresholdConditionListBox);
            this.Controls.Add(this.mediumLowerThresholdValueLabel);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.mediumUpperNumericUpDown);
            this.Controls.Add(this.highUpperNumericUpDown);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ChooseTimeThresholdsForm";
            this.Text = "Choose Time Thresholds";
            ((System.ComponentModel.ISupportInitialize)(this.highUpperNumericUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mediumUpperNumericUpDown)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.NumericUpDown highUpperNumericUpDown;
        private System.Windows.Forms.NumericUpDown mediumUpperNumericUpDown;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label mediumLowerThresholdValueLabel;
        private System.Windows.Forms.ListBox highUpperThresholdConditionListBox;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label mediumLowerThresholdConditionLabel;
        private System.Windows.Forms.ListBox mediumUpperThresholdConditionListBox;
        private System.Windows.Forms.Label routineLowerThresholdConditionLabel;
        private System.Windows.Forms.Label routineLowerThresholdValueLabel;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
    }
}