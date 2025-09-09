namespace DECS_Excel_Add_Ins
{
    partial class MergeRowsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MergeRowsForm));
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.patientDefinitionColumnsListBox = new System.Windows.Forms.ListBox();
            this.infoColumnsListBox = new System.Windows.Forms.ListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.sourceSheetListBox = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.quitButton = new System.Windows.Forms.Button();
            this.runButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dateColumnsListBox = new System.Windows.Forms.ListBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.helpButton = new System.Windows.Forms.Button();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.patientDefinitionColumnsListBox);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(23, 175);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(275, 200);
            this.groupBox3.TabIndex = 1;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Patient Defn Columns (like MRN, name)";
            // 
            // patientDefinitionColumnsListBox
            // 
            this.patientDefinitionColumnsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.patientDefinitionColumnsListBox.FormattingEnabled = true;
            this.patientDefinitionColumnsListBox.ItemHeight = 15;
            this.patientDefinitionColumnsListBox.Location = new System.Drawing.Point(15, 28);
            this.patientDefinitionColumnsListBox.Name = "patientDefinitionColumnsListBox";
            this.patientDefinitionColumnsListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.patientDefinitionColumnsListBox.Size = new System.Drawing.Size(220, 154);
            this.patientDefinitionColumnsListBox.TabIndex = 0;
            this.patientDefinitionColumnsListBox.SelectedIndexChanged += new System.EventHandler(this.PatientDefinitionColumnsListBox_SelectedIndexChanged);
            // 
            // infoColumnsListBox
            // 
            this.infoColumnsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.infoColumnsListBox.FormattingEnabled = true;
            this.infoColumnsListBox.ItemHeight = 15;
            this.infoColumnsListBox.Location = new System.Drawing.Point(13, 28);
            this.infoColumnsListBox.Name = "infoColumnsListBox";
            this.infoColumnsListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.infoColumnsListBox.Size = new System.Drawing.Size(220, 154);
            this.infoColumnsListBox.TabIndex = 0;
            this.infoColumnsListBox.SelectedIndexChanged += new System.EventHandler(this.InfoColumnsListBox_SelectedIndexChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.sourceSheetListBox);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(332, 63);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(275, 86);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Source Sheet";
            // 
            // sourceSheetListBox
            // 
            this.sourceSheetListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sourceSheetListBox.FormattingEnabled = true;
            this.sourceSheetListBox.ItemHeight = 15;
            this.sourceSheetListBox.Location = new System.Drawing.Point(13, 27);
            this.sourceSheetListBox.Name = "sourceSheetListBox";
            this.sourceSheetListBox.Size = new System.Drawing.Size(220, 34);
            this.sourceSheetListBox.TabIndex = 0;
            this.sourceSheetListBox.SelectedIndexChanged += new System.EventHandler(this.SourceSheetListBox_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(402, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(168, 25);
            this.label1.TabIndex = 7;
            this.label1.Text = "Combine Rows";
            // 
            // quitButton
            // 
            this.quitButton.BackColor = System.Drawing.Color.White;
            this.quitButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.quitButton.Location = new System.Drawing.Point(537, 504);
            this.quitButton.Name = "quitButton";
            this.quitButton.Size = new System.Drawing.Size(120, 40);
            this.quitButton.TabIndex = 12;
            this.quitButton.Text = "Quit";
            this.quitButton.UseVisualStyleBackColor = false;
            this.quitButton.Click += new System.EventHandler(this.QuitButton_Click);
            // 
            // runButton
            // 
            this.runButton.BackColor = System.Drawing.Color.White;
            this.runButton.Enabled = false;
            this.runButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.runButton.Location = new System.Drawing.Point(312, 504);
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(120, 40);
            this.runButton.TabIndex = 11;
            this.runButton.Text = "Run";
            this.runButton.UseVisualStyleBackColor = false;
            this.runButton.Click += new System.EventHandler(this.RunButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dateColumnsListBox);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(647, 175);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(275, 200);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Date Columns (like address start/end)";
            // 
            // dateColumnsListBox
            // 
            this.dateColumnsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateColumnsListBox.FormattingEnabled = true;
            this.dateColumnsListBox.ItemHeight = 15;
            this.dateColumnsListBox.Location = new System.Drawing.Point(13, 28);
            this.dateColumnsListBox.Name = "dateColumnsListBox";
            this.dateColumnsListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.dateColumnsListBox.Size = new System.Drawing.Size(220, 154);
            this.dateColumnsListBox.TabIndex = 0;
            this.dateColumnsListBox.SelectedIndexChanged += new System.EventHandler(this.DateColumnsListBox_SelectedIndexChanged);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.infoColumnsListBox);
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.Location = new System.Drawing.Point(332, 175);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(275, 200);
            this.groupBox4.TabIndex = 14;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Info Columns (like insurer, address)";
            // 
            // helpButton
            // 
            this.helpButton.BackColor = System.Drawing.Color.White;
            this.helpButton.BackgroundImage = global::DECS_Excel_Add_Ins.Properties.Resources.help_small;
            this.helpButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.helpButton.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.helpButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.helpButton.Location = new System.Drawing.Point(863, 494);
            this.helpButton.Name = "helpButton";
            this.helpButton.Size = new System.Drawing.Size(50, 50);
            this.helpButton.TabIndex = 15;
            this.helpButton.UseVisualStyleBackColor = false;
            this.helpButton.Click += new System.EventHandler(this.helpButton_Click);
            // 
            // MergeRowsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(945, 556);
            this.Controls.Add(this.helpButton);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.quitButton);
            this.Controls.Add(this.runButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MergeRowsForm";
            this.Text = "CombineRowsForm";
            this.groupBox3.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ListBox infoColumnsListBox;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ListBox sourceSheetListBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button quitButton;
        private System.Windows.Forms.Button runButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox dateColumnsListBox;
        private System.Windows.Forms.ListBox patientDefinitionColumnsListBox;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button helpButton;
    }
}