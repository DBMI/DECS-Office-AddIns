namespace DECS_Excel_Add_Ins
{
    partial class PlotSelectionForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PlotSelectionForm));
            this.label1 = new System.Windows.Forms.Label();
            this.timeColumn_1_ListBox = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.sheet_1_ListBox = new System.Windows.Forms.ListBox();
            this.nameColumn_1_ListBox = new System.Windows.Forms.ListBox();
            this.valueColumn_1_ListBox = new System.Windows.Forms.ListBox();
            this.timeColumn_2_ListBox = new System.Windows.Forms.ListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.sheet_2_ListBox = new System.Windows.Forms.ListBox();
            this.valueColumn_2_ListBox = new System.Windows.Forms.ListBox();
            this.nameColumn_2_ListBox = new System.Windows.Forms.ListBox();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(136, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(253, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select Plot Parameters";
            // 
            // timeColumn_1_ListBox
            // 
            this.timeColumn_1_ListBox.Enabled = false;
            this.timeColumn_1_ListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.timeColumn_1_ListBox.FormattingEnabled = true;
            this.timeColumn_1_ListBox.Location = new System.Drawing.Point(19, 171);
            this.timeColumn_1_ListBox.Name = "timeColumn_1_ListBox";
            this.timeColumn_1_ListBox.Size = new System.Drawing.Size(176, 95);
            this.timeColumn_1_ListBox.TabIndex = 1;
            this.timeColumn_1_ListBox.Click += new System.EventHandler(this.Time1ColumnListBox_SelectedIndexChanged);
            this.timeColumn_1_ListBox.SelectedIndexChanged += new System.EventHandler(this.Time1ColumnListBox_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.sheet_1_ListBox);
            this.groupBox1.Controls.Add(this.nameColumn_1_ListBox);
            this.groupBox1.Controls.Add(this.timeColumn_1_ListBox);
            this.groupBox1.Controls.Add(this.valueColumn_1_ListBox);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(18, 86);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(218, 531);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Sheet 1";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(16, 394);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(101, 16);
            this.label8.TabIndex = 8;
            this.label8.Text = "Name column";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(16, 273);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(100, 16);
            this.label6.TabIndex = 7;
            this.label6.Text = "Value column";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(16, 152);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(95, 16);
            this.label4.TabIndex = 5;
            this.label4.Text = "Time column";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(16, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Sheet name";
            // 
            // sheet_1_ListBox
            // 
            this.sheet_1_ListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sheet_1_ListBox.FormattingEnabled = true;
            this.sheet_1_ListBox.Location = new System.Drawing.Point(19, 50);
            this.sheet_1_ListBox.Name = "sheet_1_ListBox";
            this.sheet_1_ListBox.Size = new System.Drawing.Size(176, 95);
            this.sheet_1_ListBox.TabIndex = 2;
            this.sheet_1_ListBox.SelectedIndexChanged += new System.EventHandler(this.Sheet1ListBox_SelectedIndexChanged);
            // 
            // nameColumn_1_ListBox
            // 
            this.nameColumn_1_ListBox.Enabled = false;
            this.nameColumn_1_ListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nameColumn_1_ListBox.FormattingEnabled = true;
            this.nameColumn_1_ListBox.Location = new System.Drawing.Point(19, 413);
            this.nameColumn_1_ListBox.Name = "nameColumn_1_ListBox";
            this.nameColumn_1_ListBox.Size = new System.Drawing.Size(176, 95);
            this.nameColumn_1_ListBox.TabIndex = 1;
            this.nameColumn_1_ListBox.Click += new System.EventHandler(this.Name1ColumnListBox_SelectedIndexChanged);
            this.nameColumn_1_ListBox.SelectedIndexChanged += new System.EventHandler(this.Name1ColumnListBox_SelectedIndexChanged);
            // 
            // valueColumn_1_ListBox
            // 
            this.valueColumn_1_ListBox.Enabled = false;
            this.valueColumn_1_ListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.valueColumn_1_ListBox.FormattingEnabled = true;
            this.valueColumn_1_ListBox.Location = new System.Drawing.Point(19, 292);
            this.valueColumn_1_ListBox.Name = "valueColumn_1_ListBox";
            this.valueColumn_1_ListBox.Size = new System.Drawing.Size(176, 95);
            this.valueColumn_1_ListBox.TabIndex = 1;
            this.valueColumn_1_ListBox.Click += new System.EventHandler(this.Value1ColumnListBox_SelectedIndexChanged);
            this.valueColumn_1_ListBox.SelectedIndexChanged += new System.EventHandler(this.Value1ColumnListBox_SelectedIndexChanged);
            // 
            // timeColumn_2_ListBox
            // 
            this.timeColumn_2_ListBox.Enabled = false;
            this.timeColumn_2_ListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.timeColumn_2_ListBox.FormattingEnabled = true;
            this.timeColumn_2_ListBox.Location = new System.Drawing.Point(21, 171);
            this.timeColumn_2_ListBox.Name = "timeColumn_2_ListBox";
            this.timeColumn_2_ListBox.Size = new System.Drawing.Size(176, 95);
            this.timeColumn_2_ListBox.TabIndex = 2;
            this.timeColumn_2_ListBox.Click += new System.EventHandler(this.Time2ColumnListBox_SelectedIndexChanged);
            this.timeColumn_2_ListBox.SelectedIndexChanged += new System.EventHandler(this.Time2ColumnListBox_SelectedIndexChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.sheet_2_ListBox);
            this.groupBox2.Controls.Add(this.valueColumn_2_ListBox);
            this.groupBox2.Controls.Add(this.nameColumn_2_ListBox);
            this.groupBox2.Controls.Add(this.timeColumn_2_ListBox);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(279, 86);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(218, 531);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Sheet 2";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(21, 394);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(101, 16);
            this.label9.TabIndex = 9;
            this.label9.Text = "Name column";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(21, 273);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(100, 16);
            this.label7.TabIndex = 8;
            this.label7.Text = "Value column";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(21, 152);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(95, 16);
            this.label5.TabIndex = 7;
            this.label5.Text = "Time column";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(21, 31);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 16);
            this.label3.TabIndex = 4;
            this.label3.Text = "Sheet name";
            // 
            // sheet_2_ListBox
            // 
            this.sheet_2_ListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sheet_2_ListBox.FormattingEnabled = true;
            this.sheet_2_ListBox.Location = new System.Drawing.Point(21, 50);
            this.sheet_2_ListBox.Name = "sheet_2_ListBox";
            this.sheet_2_ListBox.Size = new System.Drawing.Size(176, 95);
            this.sheet_2_ListBox.TabIndex = 3;
            this.sheet_2_ListBox.SelectedIndexChanged += new System.EventHandler(this.Sheet2ListBox_SelectedIndexChanged);
            // 
            // valueColumn_2_ListBox
            // 
            this.valueColumn_2_ListBox.Enabled = false;
            this.valueColumn_2_ListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.valueColumn_2_ListBox.FormattingEnabled = true;
            this.valueColumn_2_ListBox.Location = new System.Drawing.Point(21, 292);
            this.valueColumn_2_ListBox.Name = "valueColumn_2_ListBox";
            this.valueColumn_2_ListBox.Size = new System.Drawing.Size(176, 95);
            this.valueColumn_2_ListBox.TabIndex = 2;
            this.valueColumn_2_ListBox.Click += new System.EventHandler(this.Value2ColumnListBox_SelectedIndexChanged);
            this.valueColumn_2_ListBox.SelectedIndexChanged += new System.EventHandler(this.Value2ColumnListBox_SelectedIndexChanged);
            // 
            // nameColumn_2_ListBox
            // 
            this.nameColumn_2_ListBox.Enabled = false;
            this.nameColumn_2_ListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nameColumn_2_ListBox.FormattingEnabled = true;
            this.nameColumn_2_ListBox.Location = new System.Drawing.Point(21, 413);
            this.nameColumn_2_ListBox.Name = "nameColumn_2_ListBox";
            this.nameColumn_2_ListBox.Size = new System.Drawing.Size(176, 95);
            this.nameColumn_2_ListBox.TabIndex = 2;
            this.nameColumn_2_ListBox.Click += new System.EventHandler(this.Name2ColumnListBox_SelectedIndexChanged);
            this.nameColumn_2_ListBox.SelectedIndexChanged += new System.EventHandler(this.Name2ColumnListBox_SelectedIndexChanged);
            // 
            // okButton
            // 
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.Location = new System.Drawing.Point(161, 639);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 5;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelButton.ForeColor = System.Drawing.Color.Black;
            this.cancelButton.Location = new System.Drawing.Point(286, 639);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 6;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // PlotSelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(546, 688);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "PlotSelectionForm";
            this.Text = "Plot Selection Form";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox timeColumn_1_ListBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox timeColumn_2_ListBox;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ListBox valueColumn_2_ListBox;
        private System.Windows.Forms.ListBox valueColumn_1_ListBox;
        private System.Windows.Forms.ListBox nameColumn_2_ListBox;
        private System.Windows.Forms.ListBox nameColumn_1_ListBox;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.ListBox sheet_1_ListBox;
        private System.Windows.Forms.ListBox sheet_2_ListBox;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
    }
}