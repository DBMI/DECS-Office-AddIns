namespace DECS_Excel_Add_Ins
{
    partial class DefineRules
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DefineRules));
            this.headline = new System.Windows.Forms.Label();
            this.sourceColumnGroupBox = new System.Windows.Forms.GroupBox();
            this.sourceColumnListBox = new System.Windows.Forms.ListBox();
            this.cleaningRulesGroupBox = new System.Windows.Forms.GroupBox();
            this.cleaningRulesPanel = new System.Windows.Forms.Panel();
            this.cleaningRulesAddButton = new System.Windows.Forms.Button();
            this.cleaningRulesReplaceLabel = new System.Windows.Forms.Label();
            this.cleaningRulesPatternsLabel = new System.Windows.Forms.Label();
            this.extractRulesGroupBox = new System.Windows.Forms.GroupBox();
            this.extractRulesPanel = new System.Windows.Forms.Panel();
            this.extractRulesAddButton = new System.Windows.Forms.Button();
            this.extractRulesnewColumnLabel = new System.Windows.Forms.Label();
            this.extractRulesPatternLabel = new System.Windows.Forms.Label();
            this.saveButton = new System.Windows.Forms.Button();
            this.discardButton = new System.Windows.Forms.Button();
            this.clearButton = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.progressBarLabel = new System.Windows.Forms.Label();
            this.sourceColumnGroupBox.SuspendLayout();
            this.cleaningRulesGroupBox.SuspendLayout();
            this.cleaningRulesPanel.SuspendLayout();
            this.extractRulesGroupBox.SuspendLayout();
            this.extractRulesPanel.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // headline
            // 
            this.headline.AutoSize = true;
            this.headline.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.headline.Location = new System.Drawing.Point(582, 27);
            this.headline.Name = "headline";
            this.headline.Size = new System.Drawing.Size(147, 25);
            this.headline.TabIndex = 0;
            this.headline.Text = "Define Rules";
            // 
            // sourceColumnGroupBox
            // 
            this.sourceColumnGroupBox.Controls.Add(this.sourceColumnListBox);
            this.sourceColumnGroupBox.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.sourceColumnGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sourceColumnGroupBox.Location = new System.Drawing.Point(31, 70);
            this.sourceColumnGroupBox.Name = "sourceColumnGroupBox";
            this.sourceColumnGroupBox.Size = new System.Drawing.Size(277, 86);
            this.sourceColumnGroupBox.TabIndex = 1;
            this.sourceColumnGroupBox.TabStop = false;
            this.sourceColumnGroupBox.Text = "Source Column";
            // 
            // sourceColumnListBox
            // 
            this.sourceColumnListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sourceColumnListBox.FormattingEnabled = true;
            this.sourceColumnListBox.ItemHeight = 15;
            this.sourceColumnListBox.Location = new System.Drawing.Point(25, 33);
            this.sourceColumnListBox.Name = "sourceColumnListBox";
            this.sourceColumnListBox.Size = new System.Drawing.Size(225, 34);
            this.sourceColumnListBox.TabIndex = 0;
            this.sourceColumnListBox.SelectedIndexChanged += new System.EventHandler(this.sourceColumnListBox_Selected);
            // 
            // cleaningRulesGroupBox
            // 
            this.cleaningRulesGroupBox.Controls.Add(this.cleaningRulesPanel);
            this.cleaningRulesGroupBox.Controls.Add(this.cleaningRulesReplaceLabel);
            this.cleaningRulesGroupBox.Controls.Add(this.cleaningRulesPatternsLabel);
            this.cleaningRulesGroupBox.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cleaningRulesGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cleaningRulesGroupBox.Location = new System.Drawing.Point(31, 162);
            this.cleaningRulesGroupBox.Name = "cleaningRulesGroupBox";
            this.cleaningRulesGroupBox.Size = new System.Drawing.Size(1350, 236);
            this.cleaningRulesGroupBox.TabIndex = 2;
            this.cleaningRulesGroupBox.TabStop = false;
            this.cleaningRulesGroupBox.Text = "Cleaning Rules";
            // 
            // cleaningRulesPanel
            // 
            this.cleaningRulesPanel.AutoScroll = true;
            this.cleaningRulesPanel.Controls.Add(this.cleaningRulesAddButton);
            this.cleaningRulesPanel.Location = new System.Drawing.Point(6, 45);
            this.cleaningRulesPanel.Name = "cleaningRulesPanel";
            this.cleaningRulesPanel.Size = new System.Drawing.Size(1340, 185);
            this.cleaningRulesPanel.TabIndex = 6;
            // 
            // cleaningRulesAddButton
            // 
            this.cleaningRulesAddButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cleaningRulesAddButton.Location = new System.Drawing.Point(1266, 23);
            this.cleaningRulesAddButton.Name = "cleaningRulesAddButton";
            this.cleaningRulesAddButton.Size = new System.Drawing.Size(40, 30);
            this.cleaningRulesAddButton.TabIndex = 5;
            this.cleaningRulesAddButton.Text = "+";
            this.cleaningRulesAddButton.UseVisualStyleBackColor = true;
            this.cleaningRulesAddButton.Click += new System.EventHandler(this.cleaningRulesAddButton_Click);
            // 
            // cleaningRulesReplaceLabel
            // 
            this.cleaningRulesReplaceLabel.AutoSize = true;
            this.cleaningRulesReplaceLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cleaningRulesReplaceLabel.Location = new System.Drawing.Point(1120, 26);
            this.cleaningRulesReplaceLabel.Name = "cleaningRulesReplaceLabel";
            this.cleaningRulesReplaceLabel.Size = new System.Drawing.Size(66, 16);
            this.cleaningRulesReplaceLabel.TabIndex = 1;
            this.cleaningRulesReplaceLabel.Text = "Replace";
            this.cleaningRulesReplaceLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // cleaningRulesPatternsLabel
            // 
            this.cleaningRulesPatternsLabel.AutoSize = true;
            this.cleaningRulesPatternsLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cleaningRulesPatternsLabel.Location = new System.Drawing.Point(500, 26);
            this.cleaningRulesPatternsLabel.Name = "cleaningRulesPatternsLabel";
            this.cleaningRulesPatternsLabel.Size = new System.Drawing.Size(56, 16);
            this.cleaningRulesPatternsLabel.TabIndex = 0;
            this.cleaningRulesPatternsLabel.Text = "Pattern";
            // 
            // extractRulesGroupBox
            // 
            this.extractRulesGroupBox.Controls.Add(this.extractRulesPanel);
            this.extractRulesGroupBox.Controls.Add(this.extractRulesnewColumnLabel);
            this.extractRulesGroupBox.Controls.Add(this.extractRulesPatternLabel);
            this.extractRulesGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.extractRulesGroupBox.Location = new System.Drawing.Point(31, 404);
            this.extractRulesGroupBox.Name = "extractRulesGroupBox";
            this.extractRulesGroupBox.Size = new System.Drawing.Size(1350, 236);
            this.extractRulesGroupBox.TabIndex = 3;
            this.extractRulesGroupBox.TabStop = false;
            this.extractRulesGroupBox.Text = "Extract Rules";
            // 
            // extractRulesPanel
            // 
            this.extractRulesPanel.AutoScroll = true;
            this.extractRulesPanel.Controls.Add(this.extractRulesAddButton);
            this.extractRulesPanel.Location = new System.Drawing.Point(6, 45);
            this.extractRulesPanel.Name = "extractRulesPanel";
            this.extractRulesPanel.Size = new System.Drawing.Size(1340, 185);
            this.extractRulesPanel.TabIndex = 4;
            // 
            // extractRulesAddButton
            // 
            this.extractRulesAddButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.extractRulesAddButton.Location = new System.Drawing.Point(1266, 25);
            this.extractRulesAddButton.Name = "extractRulesAddButton";
            this.extractRulesAddButton.Size = new System.Drawing.Size(40, 30);
            this.extractRulesAddButton.TabIndex = 8;
            this.extractRulesAddButton.Text = "+";
            this.extractRulesAddButton.UseVisualStyleBackColor = true;
            this.extractRulesAddButton.Click += new System.EventHandler(this.extractRulesAddButton_Click);
            // 
            // extractRulesnewColumnLabel
            // 
            this.extractRulesnewColumnLabel.AutoSize = true;
            this.extractRulesnewColumnLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.extractRulesnewColumnLabel.Location = new System.Drawing.Point(1110, 26);
            this.extractRulesnewColumnLabel.Name = "extractRulesnewColumnLabel";
            this.extractRulesnewColumnLabel.Size = new System.Drawing.Size(92, 16);
            this.extractRulesnewColumnLabel.TabIndex = 3;
            this.extractRulesnewColumnLabel.Text = "New Column";
            this.extractRulesnewColumnLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // extractRulesPatternLabel
            // 
            this.extractRulesPatternLabel.AutoSize = true;
            this.extractRulesPatternLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.extractRulesPatternLabel.Location = new System.Drawing.Point(500, 26);
            this.extractRulesPatternLabel.Name = "extractRulesPatternLabel";
            this.extractRulesPatternLabel.Size = new System.Drawing.Size(56, 16);
            this.extractRulesPatternLabel.TabIndex = 2;
            this.extractRulesPatternLabel.Text = "Pattern";
            // 
            // saveButton
            // 
            this.saveButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.saveButton.Location = new System.Drawing.Point(461, 662);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(120, 40);
            this.saveButton.TabIndex = 4;
            this.saveButton.Text = "Save";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // discardButton
            // 
            this.discardButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.discardButton.Location = new System.Drawing.Point(842, 662);
            this.discardButton.Name = "discardButton";
            this.discardButton.Size = new System.Drawing.Size(120, 40);
            this.discardButton.TabIndex = 5;
            this.discardButton.Text = "Discard";
            this.discardButton.UseVisualStyleBackColor = true;
            this.discardButton.Click += new System.EventHandler(this.discardButton_Click);
            // 
            // clearButton
            // 
            this.clearButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clearButton.Location = new System.Drawing.Point(662, 662);
            this.clearButton.Name = "clearButton";
            this.clearButton.Size = new System.Drawing.Size(120, 40);
            this.clearButton.TabIndex = 6;
            this.clearButton.Text = "Clear";
            this.clearButton.UseVisualStyleBackColor = true;
            this.clearButton.Click += new System.EventHandler(this.clearButton_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1414, 24);
            this.menuStrip1.TabIndex = 7;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.loadToolStripMenuItem,
            this.saveToolStripMenuItem,
            this.saveAsToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // loadToolStripMenuItem
            // 
            this.loadToolStripMenuItem.Name = "loadToolStripMenuItem";
            this.loadToolStripMenuItem.Size = new System.Drawing.Size(111, 22);
            this.loadToolStripMenuItem.Text = "Load";
            this.loadToolStripMenuItem.Click += new System.EventHandler(this.loadToolStripMenuItem_Click);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(111, 22);
            this.saveToolStripMenuItem.Text = "Save";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // saveAsToolStripMenuItem
            // 
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(111, 22);
            this.saveAsToolStripMenuItem.Text = "SaveAs";
            this.saveAsToolStripMenuItem.Click += new System.EventHandler(this.saveAsToolStripMenuItem_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(19, 51);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(239, 23);
            this.progressBar.TabIndex = 8;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.progressBarLabel);
            this.groupBox1.Controls.Add(this.progressBar);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(1104, 70);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(277, 86);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Progress";
            // 
            // progressBarLabel
            // 
            this.progressBarLabel.AutoSize = true;
            this.progressBarLabel.Location = new System.Drawing.Point(91, 22);
            this.progressBarLabel.Name = "progressBarLabel";
            this.progressBarLabel.Size = new System.Drawing.Size(97, 15);
            this.progressBarLabel.TabIndex = 9;
            this.progressBarLabel.Text = "Applying rules";
            this.progressBarLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // DefineRules
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(1414, 714);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.discardButton);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.extractRulesGroupBox);
            this.Controls.Add(this.cleaningRulesGroupBox);
            this.Controls.Add(this.sourceColumnGroupBox);
            this.Controls.Add(this.headline);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "DefineRules";
            this.Text = "Define Data Rules";
            this.sourceColumnGroupBox.ResumeLayout(false);
            this.cleaningRulesGroupBox.ResumeLayout(false);
            this.cleaningRulesGroupBox.PerformLayout();
            this.cleaningRulesPanel.ResumeLayout(false);
            this.extractRulesGroupBox.ResumeLayout(false);
            this.extractRulesGroupBox.PerformLayout();
            this.extractRulesPanel.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label headline;
        private System.Windows.Forms.GroupBox sourceColumnGroupBox;
        private System.Windows.Forms.ListBox sourceColumnListBox;
        private System.Windows.Forms.GroupBox cleaningRulesGroupBox;
        private System.Windows.Forms.Label cleaningRulesReplaceLabel;
        private System.Windows.Forms.Label cleaningRulesPatternsLabel;
        private System.Windows.Forms.GroupBox extractRulesGroupBox;
        private System.Windows.Forms.Label extractRulesnewColumnLabel;
        private System.Windows.Forms.Label extractRulesPatternLabel;
        private System.Windows.Forms.Button saveButton;
        private System.Windows.Forms.Button discardButton;
        private System.Windows.Forms.Button clearButton;
        private System.Windows.Forms.Button cleaningRulesAddButton;
        private System.Windows.Forms.Button extractRulesAddButton;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
        private System.Windows.Forms.Panel cleaningRulesPanel;
        private System.Windows.Forms.Panel extractRulesPanel;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label progressBarLabel;
    }
}