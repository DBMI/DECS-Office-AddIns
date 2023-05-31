namespace DECS_Excel_Add_Ins
{
    partial class DefineRulesForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DefineRulesForm));
            this.headline = new System.Windows.Forms.Label();
            this.sourceColumnGroupBox = new System.Windows.Forms.GroupBox();
            this.sourceColumnListBox = new System.Windows.Forms.ListBox();
            this.saveButton = new System.Windows.Forms.Button();
            this.discardButton = new System.Windows.Forms.Button();
            this.clearButton = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.runButton = new System.Windows.Forms.Button();
            this.selectedRowsGroupBox = new System.Windows.Forms.GroupBox();
            this.selectedRowsLabel = new System.Windows.Forms.Label();
            this.extractRulesPanel = new System.Windows.Forms.Panel();
            this.extractRulesAddButton = new System.Windows.Forms.Button();
            this.rulesTabControl = new System.Windows.Forms.TabControl();
            this.cleaningTabPage = new System.Windows.Forms.TabPage();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.cleaningRulesPanel = new System.Windows.Forms.Panel();
            this.cleaningRulesAddButton = new System.Windows.Forms.Button();
            this.dateFormatTabPage = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dateFormatsListBox = new System.Windows.Forms.ListBox();
            this.dateConversionEnabledCheckBox = new System.Windows.Forms.CheckBox();
            this.extractTabPage = new System.Windows.Forms.TabPage();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.toolTipBox = new System.Windows.Forms.ToolTip(this.components);
            this.sourceColumnGroupBox.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.selectedRowsGroupBox.SuspendLayout();
            this.extractRulesPanel.SuspendLayout();
            this.rulesTabControl.SuspendLayout();
            this.cleaningTabPage.SuspendLayout();
            this.cleaningRulesPanel.SuspendLayout();
            this.dateFormatTabPage.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.extractTabPage.SuspendLayout();
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
            // saveButton
            // 
            this.saveButton.BackColor = System.Drawing.Color.White;
            this.saveButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.saveButton.Location = new System.Drawing.Point(461, 662);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(120, 40);
            this.saveButton.TabIndex = 4;
            this.saveButton.Text = "Save";
            this.saveButton.UseVisualStyleBackColor = false;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // discardButton
            // 
            this.discardButton.BackColor = System.Drawing.Color.White;
            this.discardButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.discardButton.Location = new System.Drawing.Point(863, 662);
            this.discardButton.Name = "discardButton";
            this.discardButton.Size = new System.Drawing.Size(120, 40);
            this.discardButton.TabIndex = 5;
            this.discardButton.Text = "Quit";
            this.discardButton.UseVisualStyleBackColor = false;
            this.discardButton.Click += new System.EventHandler(this.discardButton_Click);
            // 
            // clearButton
            // 
            this.clearButton.BackColor = System.Drawing.Color.White;
            this.clearButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clearButton.Location = new System.Drawing.Point(662, 662);
            this.clearButton.Name = "clearButton";
            this.clearButton.Size = new System.Drawing.Size(120, 40);
            this.clearButton.TabIndex = 6;
            this.clearButton.Text = "Clear";
            this.clearButton.UseVisualStyleBackColor = false;
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
            // runButton
            // 
            this.runButton.BackColor = System.Drawing.Color.White;
            this.runButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.runButton.ForeColor = System.Drawing.Color.DarkBlue;
            this.runButton.Location = new System.Drawing.Point(662, 103);
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(120, 40);
            this.runButton.TabIndex = 8;
            this.runButton.Text = "Run";
            this.runButton.UseVisualStyleBackColor = false;
            this.runButton.Click += new System.EventHandler(this.runButton_Click);
            // 
            // selectedRowsGroupBox
            // 
            this.selectedRowsGroupBox.Controls.Add(this.selectedRowsLabel);
            this.selectedRowsGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.selectedRowsGroupBox.Location = new System.Drawing.Point(1104, 70);
            this.selectedRowsGroupBox.Name = "selectedRowsGroupBox";
            this.selectedRowsGroupBox.Size = new System.Drawing.Size(277, 86);
            this.selectedRowsGroupBox.TabIndex = 9;
            this.selectedRowsGroupBox.TabStop = false;
            this.selectedRowsGroupBox.Text = "Rows Selected for Processing";
            // 
            // selectedRowsLabel
            // 
            this.selectedRowsLabel.AutoSize = true;
            this.selectedRowsLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.selectedRowsLabel.Location = new System.Drawing.Point(7, 33);
            this.selectedRowsLabel.Name = "selectedRowsLabel";
            this.selectedRowsLabel.Size = new System.Drawing.Size(0, 15);
            this.selectedRowsLabel.TabIndex = 0;
            // 
            // extractRulesPanel
            // 
            this.extractRulesPanel.AutoScroll = true;
            this.extractRulesPanel.Controls.Add(this.extractRulesAddButton);
            this.extractRulesPanel.Location = new System.Drawing.Point(6, 38);
            this.extractRulesPanel.Name = "extractRulesPanel";
            this.extractRulesPanel.Size = new System.Drawing.Size(1373, 412);
            this.extractRulesPanel.TabIndex = 4;
            // 
            // extractRulesAddButton
            // 
            this.extractRulesAddButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.extractRulesAddButton.Location = new System.Drawing.Point(1266, 23);
            this.extractRulesAddButton.Name = "extractRulesAddButton";
            this.extractRulesAddButton.Size = new System.Drawing.Size(40, 30);
            this.extractRulesAddButton.TabIndex = 8;
            this.extractRulesAddButton.Text = "+";
            this.extractRulesAddButton.UseVisualStyleBackColor = true;
            this.extractRulesAddButton.Click += new System.EventHandler(this.extractRulesAddButton_Click);
            // 
            // rulesTabControl
            // 
            this.rulesTabControl.Controls.Add(this.cleaningTabPage);
            this.rulesTabControl.Controls.Add(this.dateFormatTabPage);
            this.rulesTabControl.Controls.Add(this.extractTabPage);
            this.rulesTabControl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rulesTabControl.Location = new System.Drawing.Point(12, 172);
            this.rulesTabControl.Name = "rulesTabControl";
            this.rulesTabControl.SelectedIndex = 0;
            this.rulesTabControl.Size = new System.Drawing.Size(1393, 484);
            this.rulesTabControl.TabIndex = 10;
            // 
            // cleaningTabPage
            // 
            this.cleaningTabPage.Controls.Add(this.label6);
            this.cleaningTabPage.Controls.Add(this.label4);
            this.cleaningTabPage.Controls.Add(this.label1);
            this.cleaningTabPage.Controls.Add(this.cleaningRulesPanel);
            this.cleaningTabPage.Location = new System.Drawing.Point(4, 24);
            this.cleaningTabPage.Name = "cleaningTabPage";
            this.cleaningTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.cleaningTabPage.Size = new System.Drawing.Size(1385, 456);
            this.cleaningTabPage.TabIndex = 0;
            this.cleaningTabPage.Text = "Cleaning";
            this.cleaningTabPage.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(1120, 20);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(60, 15);
            this.label6.TabIndex = 11;
            this.label6.Text = "Replace";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(500, 20);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 15);
            this.label4.TabIndex = 10;
            this.label4.Text = "Pattern";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(4, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Enabled";
            // 
            // cleaningRulesPanel
            // 
            this.cleaningRulesPanel.AutoScroll = true;
            this.cleaningRulesPanel.Controls.Add(this.cleaningRulesAddButton);
            this.cleaningRulesPanel.Location = new System.Drawing.Point(6, 38);
            this.cleaningRulesPanel.Name = "cleaningRulesPanel";
            this.cleaningRulesPanel.Size = new System.Drawing.Size(1373, 412);
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
            // dateFormatTabPage
            // 
            this.dateFormatTabPage.Controls.Add(this.groupBox1);
            this.dateFormatTabPage.Controls.Add(this.dateConversionEnabledCheckBox);
            this.dateFormatTabPage.Location = new System.Drawing.Point(4, 24);
            this.dateFormatTabPage.Name = "dateFormatTabPage";
            this.dateFormatTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.dateFormatTabPage.Size = new System.Drawing.Size(1385, 456);
            this.dateFormatTabPage.TabIndex = 2;
            this.dateFormatTabPage.Text = "Date Format";
            this.dateFormatTabPage.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dateFormatsListBox);
            this.groupBox1.Location = new System.Drawing.Point(40, 93);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(285, 198);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Desired Date Format";
            // 
            // dateFormatsListBox
            // 
            this.dateFormatsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateFormatsListBox.FormattingEnabled = true;
            this.dateFormatsListBox.ItemHeight = 15;
            this.dateFormatsListBox.Location = new System.Drawing.Point(14, 20);
            this.dateFormatsListBox.Name = "dateFormatsListBox";
            this.dateFormatsListBox.Size = new System.Drawing.Size(253, 169);
            this.dateFormatsListBox.TabIndex = 1;
            this.dateFormatsListBox.SelectedIndexChanged += new System.EventHandler(this.dateFormatsListBox_SelectedIndexChanged);
            // 
            // dateConversionEnabledCheckBox
            // 
            this.dateConversionEnabledCheckBox.AutoSize = true;
            this.dateConversionEnabledCheckBox.Location = new System.Drawing.Point(40, 59);
            this.dateConversionEnabledCheckBox.Name = "dateConversionEnabledCheckBox";
            this.dateConversionEnabledCheckBox.Size = new System.Drawing.Size(79, 19);
            this.dateConversionEnabledCheckBox.TabIndex = 0;
            this.dateConversionEnabledCheckBox.Text = "Enabled";
            this.dateConversionEnabledCheckBox.UseVisualStyleBackColor = true;
            this.dateConversionEnabledCheckBox.Click += new System.EventHandler(this.dateConversionEnabledCheckBox_Click);
            // 
            // extractTabPage
            // 
            this.extractTabPage.Controls.Add(this.label5);
            this.extractTabPage.Controls.Add(this.label3);
            this.extractTabPage.Controls.Add(this.label2);
            this.extractTabPage.Controls.Add(this.extractRulesPanel);
            this.extractTabPage.Location = new System.Drawing.Point(4, 24);
            this.extractTabPage.Name = "extractTabPage";
            this.extractTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.extractTabPage.Size = new System.Drawing.Size(1385, 456);
            this.extractTabPage.TabIndex = 1;
            this.extractTabPage.Text = "Extract";
            this.extractTabPage.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(1110, 20);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(88, 15);
            this.label5.TabIndex = 10;
            this.label5.Text = "New Column";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(500, 20);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 15);
            this.label3.TabIndex = 9;
            this.label3.Text = "Pattern";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(4, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Enabled";
            // 
            // toolTipBox
            // 
            this.toolTipBox.AutoPopDelay = 1000;
            this.toolTipBox.InitialDelay = 250;
            this.toolTipBox.ReshowDelay = 100;
            // 
            // DefineRulesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(1414, 714);
            this.Controls.Add(this.selectedRowsGroupBox);
            this.Controls.Add(this.runButton);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.discardButton);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.sourceColumnGroupBox);
            this.Controls.Add(this.headline);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.rulesTabControl);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "DefineRulesForm";
            this.Text = "Define Data Rules";
            this.sourceColumnGroupBox.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.selectedRowsGroupBox.ResumeLayout(false);
            this.selectedRowsGroupBox.PerformLayout();
            this.extractRulesPanel.ResumeLayout(false);
            this.rulesTabControl.ResumeLayout(false);
            this.cleaningTabPage.ResumeLayout(false);
            this.cleaningTabPage.PerformLayout();
            this.cleaningRulesPanel.ResumeLayout(false);
            this.dateFormatTabPage.ResumeLayout(false);
            this.dateFormatTabPage.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.extractTabPage.ResumeLayout(false);
            this.extractTabPage.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label headline;
        private System.Windows.Forms.GroupBox sourceColumnGroupBox;
        private System.Windows.Forms.ListBox sourceColumnListBox;
        private System.Windows.Forms.Button saveButton;
        private System.Windows.Forms.Button discardButton;
        private System.Windows.Forms.Button clearButton;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
        private System.Windows.Forms.Button runButton;
        private System.Windows.Forms.GroupBox selectedRowsGroupBox;
        private System.Windows.Forms.Label selectedRowsLabel;
        private System.Windows.Forms.Panel extractRulesPanel;
        private System.Windows.Forms.Button extractRulesAddButton;
        private System.Windows.Forms.TabControl rulesTabControl;
        private System.Windows.Forms.TabPage cleaningTabPage;
        private System.Windows.Forms.Panel cleaningRulesPanel;
        private System.Windows.Forms.Button cleaningRulesAddButton;
        private System.Windows.Forms.TabPage dateFormatTabPage;
        private System.Windows.Forms.TabPage extractTabPage;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox dateFormatsListBox;
        private System.Windows.Forms.CheckBox dateConversionEnabledCheckBox;
        private System.Windows.Forms.ToolTip toolTipBox;
    }
}