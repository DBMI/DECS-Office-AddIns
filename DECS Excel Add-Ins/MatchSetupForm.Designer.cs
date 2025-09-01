namespace DECS_Excel_Add_Ins
{
    partial class MatchSetupForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MatchSetupForm));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.sourceSheetsListBox = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.targetSheetsListBox = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.sourceNameColumnListBox = new System.Windows.Forms.ListBox();
            this.label6 = new System.Windows.Forms.Label();
            this.idColumnListBox = new System.Windows.Forms.ListBox();
            this.label7 = new System.Windows.Forms.Label();
            this.targetNameColumnListBox = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(172, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Source";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label2.Location = new System.Drawing.Point(539, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 25);
            this.label2.TabIndex = 1;
            this.label2.Text = "Target";
            // 
            // okButton
            // 
            this.okButton.Enabled = false;
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.Location = new System.Drawing.Point(258, 353);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(120, 44);
            this.okButton.TabIndex = 2;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelButton.Location = new System.Drawing.Point(458, 353);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(120, 44);
            this.cancelButton.TabIndex = 3;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // sourceSheetsListBox
            // 
            this.sourceSheetsListBox.Enabled = false;
            this.sourceSheetsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sourceSheetsListBox.FormattingEnabled = true;
            this.sourceSheetsListBox.ItemHeight = 16;
            this.sourceSheetsListBox.Location = new System.Drawing.Point(113, 123);
            this.sourceSheetsListBox.Name = "sourceSheetsListBox";
            this.sourceSheetsListBox.Size = new System.Drawing.Size(161, 36);
            this.sourceSheetsListBox.TabIndex = 4;
            this.sourceSheetsListBox.SelectedIndexChanged += new System.EventHandler(this.sourceSheetsListBox_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(113, 103);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 16);
            this.label3.TabIndex = 5;
            this.label3.Text = "Sheet";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(491, 103);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(47, 16);
            this.label4.TabIndex = 7;
            this.label4.Text = "Sheet";
            // 
            // targetSheetsListBox
            // 
            this.targetSheetsListBox.Enabled = false;
            this.targetSheetsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.targetSheetsListBox.FormattingEnabled = true;
            this.targetSheetsListBox.ItemHeight = 16;
            this.targetSheetsListBox.Location = new System.Drawing.Point(491, 123);
            this.targetSheetsListBox.Name = "targetSheetsListBox";
            this.targetSheetsListBox.Size = new System.Drawing.Size(161, 36);
            this.targetSheetsListBox.TabIndex = 6;
            this.targetSheetsListBox.SelectedIndexChanged += new System.EventHandler(this.targetSheetsListBox_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(28, 212);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(103, 16);
            this.label5.TabIndex = 9;
            this.label5.Text = "Name Column";
            // 
            // sourceNameColumnListBox
            // 
            this.sourceNameColumnListBox.Enabled = false;
            this.sourceNameColumnListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sourceNameColumnListBox.FormattingEnabled = true;
            this.sourceNameColumnListBox.ItemHeight = 16;
            this.sourceNameColumnListBox.Location = new System.Drawing.Point(28, 232);
            this.sourceNameColumnListBox.Name = "sourceNameColumnListBox";
            this.sourceNameColumnListBox.Size = new System.Drawing.Size(161, 68);
            this.sourceNameColumnListBox.TabIndex = 8;
            this.sourceNameColumnListBox.SelectedIndexChanged += new System.EventHandler(this.EnableWhenReady);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(231, 212);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(75, 16);
            this.label6.TabIndex = 11;
            this.label6.Text = "ID column";
            // 
            // idColumnListBox
            // 
            this.idColumnListBox.Enabled = false;
            this.idColumnListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.idColumnListBox.FormattingEnabled = true;
            this.idColumnListBox.ItemHeight = 16;
            this.idColumnListBox.Location = new System.Drawing.Point(231, 232);
            this.idColumnListBox.Name = "idColumnListBox";
            this.idColumnListBox.Size = new System.Drawing.Size(161, 68);
            this.idColumnListBox.TabIndex = 10;
            this.idColumnListBox.SelectedIndexChanged += new System.EventHandler(this.EnableWhenReady);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(494, 212);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(103, 16);
            this.label7.TabIndex = 13;
            this.label7.Text = "Name Column";
            // 
            // targetNameColumnListBox
            // 
            this.targetNameColumnListBox.Enabled = false;
            this.targetNameColumnListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.targetNameColumnListBox.FormattingEnabled = true;
            this.targetNameColumnListBox.ItemHeight = 16;
            this.targetNameColumnListBox.Location = new System.Drawing.Point(494, 232);
            this.targetNameColumnListBox.Name = "targetNameColumnListBox";
            this.targetNameColumnListBox.Size = new System.Drawing.Size(161, 68);
            this.targetNameColumnListBox.TabIndex = 12;
            this.targetNameColumnListBox.SelectedIndexChanged += new System.EventHandler(this.EnableWhenReady);
            // 
            // MatchSetupForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.targetNameColumnListBox);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.idColumnListBox);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.sourceNameColumnListBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.targetSheetsListBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.sourceSheetsListBox);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MatchSetupForm";
            this.Text = "Match Setup Form";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.ListBox sourceSheetsListBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListBox targetSheetsListBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ListBox sourceNameColumnListBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ListBox idColumnListBox;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ListBox targetNameColumnListBox;
    }
}