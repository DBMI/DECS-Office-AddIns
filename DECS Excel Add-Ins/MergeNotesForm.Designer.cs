namespace DECS_Excel_Add_Ins
{
    partial class MergeNotesForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MergeNotesForm));
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.dateColumnListBox = new System.Windows.Forms.ListBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.sourceColumnsListBox = new System.Windows.Forms.ListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.sourceSheetListBox = new System.Windows.Forms.ListBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.targetSheetListBox = new System.Windows.Forms.ListBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.indexColumnListBox = new System.Windows.Forms.ListBox();
            this.runButton = new System.Windows.Forms.Button();
            this.quitButton = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(403, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(146, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Merge Notes";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox5);
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(31, 73);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(288, 357);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Source";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.dateColumnListBox);
            this.groupBox5.Location = new System.Drawing.Point(15, 140);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(250, 72);
            this.groupBox5.TabIndex = 2;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Date Column";
            // 
            // dateColumnListBox
            // 
            this.dateColumnListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateColumnListBox.FormattingEnabled = true;
            this.dateColumnListBox.ItemHeight = 15;
            this.dateColumnListBox.Location = new System.Drawing.Point(12, 21);
            this.dateColumnListBox.Name = "dateColumnListBox";
            this.dateColumnListBox.Size = new System.Drawing.Size(220, 34);
            this.dateColumnListBox.TabIndex = 0;
            this.dateColumnListBox.SelectedIndexChanged += new System.EventHandler(this.DateColumnListBox_SelectedIndexChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.sourceColumnsListBox);
            this.groupBox3.Location = new System.Drawing.Point(15, 218);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(250, 123);
            this.groupBox3.TabIndex = 1;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Data Columns";
            // 
            // sourceColumnsListBox
            // 
            this.sourceColumnsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sourceColumnsListBox.FormattingEnabled = true;
            this.sourceColumnsListBox.ItemHeight = 15;
            this.sourceColumnsListBox.Location = new System.Drawing.Point(12, 31);
            this.sourceColumnsListBox.Name = "sourceColumnsListBox";
            this.sourceColumnsListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.sourceColumnsListBox.Size = new System.Drawing.Size(220, 64);
            this.sourceColumnsListBox.TabIndex = 0;
            this.sourceColumnsListBox.SelectedIndexChanged += new System.EventHandler(this.SourceColumnsListBox_SelectedIndexChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.sourceSheetListBox);
            this.groupBox2.Location = new System.Drawing.Point(15, 37);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(250, 86);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Sheet";
            // 
            // sourceSheetListBox
            // 
            this.sourceSheetListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sourceSheetListBox.FormattingEnabled = true;
            this.sourceSheetListBox.ItemHeight = 15;
            this.sourceSheetListBox.Location = new System.Drawing.Point(12, 30);
            this.sourceSheetListBox.Name = "sourceSheetListBox";
            this.sourceSheetListBox.Size = new System.Drawing.Size(220, 34);
            this.sourceSheetListBox.TabIndex = 0;
            this.sourceSheetListBox.SelectedIndexChanged += new System.EventHandler(this.SourceSheetListBox_SelectedIndexChanged);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.groupBox6);
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.Location = new System.Drawing.Point(627, 73);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(288, 357);
            this.groupBox4.TabIndex = 2;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Target";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.targetSheetListBox);
            this.groupBox6.Location = new System.Drawing.Point(15, 37);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(250, 86);
            this.groupBox6.TabIndex = 0;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Sheet";
            // 
            // targetSheetListBox
            // 
            this.targetSheetListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.targetSheetListBox.FormattingEnabled = true;
            this.targetSheetListBox.ItemHeight = 15;
            this.targetSheetListBox.Location = new System.Drawing.Point(16, 30);
            this.targetSheetListBox.Name = "targetSheetListBox";
            this.targetSheetListBox.Size = new System.Drawing.Size(220, 34);
            this.targetSheetListBox.TabIndex = 0;
            this.targetSheetListBox.SelectedIndexChanged += new System.EventHandler(this.TargetSheetListBox_SelectedIndexChanged);
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.indexColumnListBox);
            this.groupBox7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox7.Location = new System.Drawing.Point(347, 182);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(250, 86);
            this.groupBox7.TabIndex = 3;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Index Column";
            // 
            // indexColumnListBox
            // 
            this.indexColumnListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.indexColumnListBox.FormattingEnabled = true;
            this.indexColumnListBox.ItemHeight = 15;
            this.indexColumnListBox.Location = new System.Drawing.Point(12, 31);
            this.indexColumnListBox.Name = "indexColumnListBox";
            this.indexColumnListBox.Size = new System.Drawing.Size(220, 34);
            this.indexColumnListBox.TabIndex = 0;
            this.indexColumnListBox.SelectedIndexChanged += new System.EventHandler(this.IndexColumnListBox_SelectedIndexChanged);
            // 
            // runButton
            // 
            this.runButton.BackColor = System.Drawing.Color.White;
            this.runButton.Enabled = false;
            this.runButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.runButton.Location = new System.Drawing.Point(274, 467);
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(120, 40);
            this.runButton.TabIndex = 5;
            this.runButton.Text = "Run";
            this.runButton.UseVisualStyleBackColor = false;
            this.runButton.Click += new System.EventHandler(this.RunButton_Click);
            // 
            // quitButton
            // 
            this.quitButton.BackColor = System.Drawing.Color.White;
            this.quitButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.quitButton.Location = new System.Drawing.Point(539, 467);
            this.quitButton.Name = "quitButton";
            this.quitButton.Size = new System.Drawing.Size(120, 40);
            this.quitButton.TabIndex = 6;
            this.quitButton.Text = "Quit";
            this.quitButton.UseVisualStyleBackColor = false;
            this.quitButton.Click += new System.EventHandler(this.QuitButton_Click);
            // 
            // MergeNotesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(945, 556);
            this.Controls.Add(this.quitButton);
            this.Controls.Add(this.runButton);
            this.Controls.Add(this.groupBox7);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MergeNotesForm";
            this.Text = "Merge Notes";
            this.groupBox1.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ListBox sourceColumnsListBox;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ListBox sourceSheetListBox;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.ListBox targetSheetListBox;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.ListBox indexColumnListBox;
        private System.Windows.Forms.Button runButton;
        private System.Windows.Forms.Button quitButton;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.ListBox dateColumnListBox;
    }
}