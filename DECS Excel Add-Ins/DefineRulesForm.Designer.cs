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
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DefineRulesForm));
            headline = new System.Windows.Forms.Label();
            sourceColumnGroupBox = new System.Windows.Forms.GroupBox();
            sourceColumnListBox = new System.Windows.Forms.ListBox();
            saveButton = new System.Windows.Forms.Button();
            discardButton = new System.Windows.Forms.Button();
            clearButton = new System.Windows.Forms.Button();
            menuStrip1 = new System.Windows.Forms.MenuStrip();
            fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            loadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            runButton = new System.Windows.Forms.Button();
            selectedRowsGroupBox = new System.Windows.Forms.GroupBox();
            selectedRowsLabel = new System.Windows.Forms.Label();
            extractRulesPanel = new System.Windows.Forms.Panel();
            extractRulesAddButton = new System.Windows.Forms.Button();
            rulesTabControl = new System.Windows.Forms.TabControl();
            cleaningTabPage = new System.Windows.Forms.TabPage();
            label6 = new System.Windows.Forms.Label();
            label4 = new System.Windows.Forms.Label();
            label1 = new System.Windows.Forms.Label();
            cleaningRulesPanel = new System.Windows.Forms.Panel();
            cleaningRulesAddButton = new System.Windows.Forms.Button();
            dateFormatTabPage = new System.Windows.Forms.TabPage();
            groupBox1 = new System.Windows.Forms.GroupBox();
            dateFormatsListBox = new System.Windows.Forms.ListBox();
            dateConversionEnabledCheckBox = new System.Windows.Forms.CheckBox();
            extractTabPage = new System.Windows.Forms.TabPage();
            label5 = new System.Windows.Forms.Label();
            label3 = new System.Windows.Forms.Label();
            label2 = new System.Windows.Forms.Label();
            toolTipBox = new System.Windows.Forms.ToolTip(components);
            label7 = new System.Windows.Forms.Label();
            label8 = new System.Windows.Forms.Label();
            sourceColumnGroupBox.SuspendLayout();
            menuStrip1.SuspendLayout();
            selectedRowsGroupBox.SuspendLayout();
            extractRulesPanel.SuspendLayout();
            rulesTabControl.SuspendLayout();
            cleaningTabPage.SuspendLayout();
            cleaningRulesPanel.SuspendLayout();
            dateFormatTabPage.SuspendLayout();
            groupBox1.SuspendLayout();
            extractTabPage.SuspendLayout();
            SuspendLayout();
            // 
            // headline
            // 
            headline.AutoSize = true;
            headline.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            headline.Location = new System.Drawing.Point(582, 27);
            headline.Name = "headline";
            headline.Size = new System.Drawing.Size(147, 25);
            headline.TabIndex = 0;
            headline.Text = "Define Rules";
            // 
            // sourceColumnGroupBox
            // 
            sourceColumnGroupBox.Controls.Add(sourceColumnListBox);
            sourceColumnGroupBox.FlatStyle = System.Windows.Forms.FlatStyle.System;
            sourceColumnGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            sourceColumnGroupBox.Location = new System.Drawing.Point(31, 70);
            sourceColumnGroupBox.Name = "sourceColumnGroupBox";
            sourceColumnGroupBox.Size = new System.Drawing.Size(277, 86);
            sourceColumnGroupBox.TabIndex = 1;
            sourceColumnGroupBox.TabStop = false;
            sourceColumnGroupBox.Text = "Source Column";
            // 
            // sourceColumnListBox
            // 
            sourceColumnListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            sourceColumnListBox.FormattingEnabled = true;
            sourceColumnListBox.ItemHeight = 15;
            sourceColumnListBox.Location = new System.Drawing.Point(25, 33);
            sourceColumnListBox.Name = "sourceColumnListBox";
            sourceColumnListBox.Size = new System.Drawing.Size(225, 34);
            sourceColumnListBox.TabIndex = 0;
            sourceColumnListBox.SelectedIndexChanged += new System.EventHandler(sourceColumnListBox_Selected);
            // 
            // saveButton
            // 
            saveButton.BackColor = System.Drawing.Color.White;
            saveButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            saveButton.Location = new System.Drawing.Point(461, 662);
            saveButton.Name = "saveButton";
            saveButton.Size = new System.Drawing.Size(120, 40);
            saveButton.TabIndex = 4;
            saveButton.Text = "Save";
            saveButton.UseVisualStyleBackColor = false;
            saveButton.Click += new System.EventHandler(saveButton_Click);
            // 
            // discardButton
            // 
            discardButton.BackColor = System.Drawing.Color.White;
            discardButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            discardButton.Location = new System.Drawing.Point(863, 662);
            discardButton.Name = "discardButton";
            discardButton.Size = new System.Drawing.Size(120, 40);
            discardButton.TabIndex = 5;
            discardButton.Text = "Quit";
            discardButton.UseVisualStyleBackColor = false;
            discardButton.Click += new System.EventHandler(discardButton_Click);
            // 
            // clearButton
            // 
            clearButton.BackColor = System.Drawing.Color.White;
            clearButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            clearButton.Location = new System.Drawing.Point(662, 662);
            clearButton.Name = "clearButton";
            clearButton.Size = new System.Drawing.Size(120, 40);
            clearButton.TabIndex = 6;
            clearButton.Text = "Clear";
            clearButton.UseVisualStyleBackColor = false;
            clearButton.Click += new System.EventHandler(clearButton_Click);
            // 
            // menuStrip1
            // 
            menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            fileToolStripMenuItem});
            menuStrip1.Location = new System.Drawing.Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new System.Drawing.Size(1414, 24);
            menuStrip1.TabIndex = 7;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            loadToolStripMenuItem,
            saveToolStripMenuItem,
            saveAsToolStripMenuItem});
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            fileToolStripMenuItem.Text = "File";
            // 
            // loadToolStripMenuItem
            // 
            loadToolStripMenuItem.Name = "loadToolStripMenuItem";
            loadToolStripMenuItem.Size = new System.Drawing.Size(111, 22);
            loadToolStripMenuItem.Text = "Load";
            loadToolStripMenuItem.Click += new System.EventHandler(loadToolStripMenuItem_Click);
            // 
            // saveToolStripMenuItem
            // 
            saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            saveToolStripMenuItem.Size = new System.Drawing.Size(111, 22);
            saveToolStripMenuItem.Text = "Save";
            saveToolStripMenuItem.Click += new System.EventHandler(saveToolStripMenuItem_Click);
            // 
            // saveAsToolStripMenuItem
            // 
            saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            saveAsToolStripMenuItem.Size = new System.Drawing.Size(111, 22);
            saveAsToolStripMenuItem.Text = "SaveAs";
            saveAsToolStripMenuItem.Click += new System.EventHandler(saveAsToolStripMenuItem_Click);
            // 
            // runButton
            // 
            runButton.BackColor = System.Drawing.Color.White;
            runButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            runButton.ForeColor = System.Drawing.Color.DarkBlue;
            runButton.Location = new System.Drawing.Point(662, 103);
            runButton.Name = "runButton";
            runButton.Size = new System.Drawing.Size(120, 40);
            runButton.TabIndex = 8;
            runButton.Text = "Run";
            runButton.UseVisualStyleBackColor = false;
            runButton.Click += new System.EventHandler(runButton_Click);
            // 
            // selectedRowsGroupBox
            // 
            selectedRowsGroupBox.Controls.Add(selectedRowsLabel);
            selectedRowsGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            selectedRowsGroupBox.Location = new System.Drawing.Point(1104, 70);
            selectedRowsGroupBox.Name = "selectedRowsGroupBox";
            selectedRowsGroupBox.Size = new System.Drawing.Size(277, 86);
            selectedRowsGroupBox.TabIndex = 9;
            selectedRowsGroupBox.TabStop = false;
            selectedRowsGroupBox.Text = "Rows Selected for Processing";
            // 
            // selectedRowsLabel
            // 
            selectedRowsLabel.AutoSize = true;
            selectedRowsLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            selectedRowsLabel.Location = new System.Drawing.Point(7, 33);
            selectedRowsLabel.Name = "selectedRowsLabel";
            selectedRowsLabel.Size = new System.Drawing.Size(0, 15);
            selectedRowsLabel.TabIndex = 0;
            // 
            // extractRulesPanel
            // 
            extractRulesPanel.AutoScroll = true;
            extractRulesPanel.Controls.Add(extractRulesAddButton);
            extractRulesPanel.Location = new System.Drawing.Point(6, 38);
            extractRulesPanel.Name = "extractRulesPanel";
            extractRulesPanel.Size = new System.Drawing.Size(1373, 412);
            extractRulesPanel.TabIndex = 4;
            // 
            // extractRulesAddButton
            // 
            extractRulesAddButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            extractRulesAddButton.Location = new System.Drawing.Point(1266, 23);
            extractRulesAddButton.Name = "extractRulesAddButton";
            extractRulesAddButton.Size = new System.Drawing.Size(40, 30);
            extractRulesAddButton.TabIndex = 8;
            extractRulesAddButton.Text = "+";
            extractRulesAddButton.UseVisualStyleBackColor = true;
            extractRulesAddButton.Click += new System.EventHandler(extractRulesAddButton_Click);
            // 
            // rulesTabControl
            // 
            rulesTabControl.Controls.Add(cleaningTabPage);
            rulesTabControl.Controls.Add(dateFormatTabPage);
            rulesTabControl.Controls.Add(extractTabPage);
            rulesTabControl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            rulesTabControl.Location = new System.Drawing.Point(12, 172);
            rulesTabControl.Name = "rulesTabControl";
            rulesTabControl.SelectedIndex = 0;
            rulesTabControl.Size = new System.Drawing.Size(1393, 484);
            rulesTabControl.TabIndex = 10;
            // 
            // cleaningTabPage
            // 
            cleaningTabPage.Controls.Add(label7);
            cleaningTabPage.Controls.Add(label6);
            cleaningTabPage.Controls.Add(label4);
            cleaningTabPage.Controls.Add(label1);
            cleaningTabPage.Controls.Add(cleaningRulesPanel);
            cleaningTabPage.Location = new System.Drawing.Point(4, 24);
            cleaningTabPage.Name = "cleaningTabPage";
            cleaningTabPage.Padding = new System.Windows.Forms.Padding(3);
            cleaningTabPage.Size = new System.Drawing.Size(1385, 456);
            cleaningTabPage.TabIndex = 0;
            cleaningTabPage.Text = "Cleaning";
            cleaningTabPage.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new System.Drawing.Point(1125, 20);
            label6.Name = "label6";
            label6.Size = new System.Drawing.Size(60, 15);
            label6.TabIndex = 11;
            label6.Text = "Replace";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new System.Drawing.Point(570, 20);
            label4.Name = "label4";
            label4.Size = new System.Drawing.Size(53, 15);
            label4.TabIndex = 10;
            label4.Text = "Pattern";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label1.Location = new System.Drawing.Point(1, 20);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(60, 15);
            label1.TabIndex = 7;
            label1.Text = "Enabled";
            // 
            // cleaningRulesPanel
            // 
            cleaningRulesPanel.AutoScroll = true;
            cleaningRulesPanel.Controls.Add(cleaningRulesAddButton);
            cleaningRulesPanel.Location = new System.Drawing.Point(6, 38);
            cleaningRulesPanel.Name = "cleaningRulesPanel";
            cleaningRulesPanel.Size = new System.Drawing.Size(1373, 412);
            cleaningRulesPanel.TabIndex = 6;
            // 
            // cleaningRulesAddButton
            // 
            cleaningRulesAddButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            cleaningRulesAddButton.Location = new System.Drawing.Point(1266, 23);
            cleaningRulesAddButton.Name = "cleaningRulesAddButton";
            cleaningRulesAddButton.Size = new System.Drawing.Size(40, 30);
            cleaningRulesAddButton.TabIndex = 5;
            cleaningRulesAddButton.Text = "+";
            cleaningRulesAddButton.UseVisualStyleBackColor = true;
            cleaningRulesAddButton.Click += new System.EventHandler(cleaningRulesAddButton_Click);
            // 
            // dateFormatTabPage
            // 
            dateFormatTabPage.Controls.Add(groupBox1);
            dateFormatTabPage.Controls.Add(dateConversionEnabledCheckBox);
            dateFormatTabPage.Location = new System.Drawing.Point(4, 24);
            dateFormatTabPage.Name = "dateFormatTabPage";
            dateFormatTabPage.Padding = new System.Windows.Forms.Padding(3);
            dateFormatTabPage.Size = new System.Drawing.Size(1385, 456);
            dateFormatTabPage.TabIndex = 2;
            dateFormatTabPage.Text = "Date Format";
            dateFormatTabPage.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(dateFormatsListBox);
            groupBox1.Location = new System.Drawing.Point(40, 93);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(285, 198);
            groupBox1.TabIndex = 2;
            groupBox1.TabStop = false;
            groupBox1.Text = "Desired Date Format";
            // 
            // dateFormatsListBox
            // 
            dateFormatsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dateFormatsListBox.FormattingEnabled = true;
            dateFormatsListBox.ItemHeight = 15;
            dateFormatsListBox.Location = new System.Drawing.Point(14, 20);
            dateFormatsListBox.Name = "dateFormatsListBox";
            dateFormatsListBox.Size = new System.Drawing.Size(253, 169);
            dateFormatsListBox.TabIndex = 1;
            dateFormatsListBox.SelectedIndexChanged += new System.EventHandler(dateFormatsListBox_SelectedIndexChanged);
            // 
            // dateConversionEnabledCheckBox
            // 
            dateConversionEnabledCheckBox.AutoSize = true;
            dateConversionEnabledCheckBox.Location = new System.Drawing.Point(40, 59);
            dateConversionEnabledCheckBox.Name = "dateConversionEnabledCheckBox";
            dateConversionEnabledCheckBox.Size = new System.Drawing.Size(79, 19);
            dateConversionEnabledCheckBox.TabIndex = 0;
            dateConversionEnabledCheckBox.Text = "Enabled";
            dateConversionEnabledCheckBox.UseVisualStyleBackColor = true;
            dateConversionEnabledCheckBox.Click += new System.EventHandler(dateConversionEnabledCheckBox_Click);
            // 
            // extractTabPage
            // 
            extractTabPage.Controls.Add(label8);
            extractTabPage.Controls.Add(label5);
            extractTabPage.Controls.Add(label3);
            extractTabPage.Controls.Add(label2);
            extractTabPage.Controls.Add(extractRulesPanel);
            extractTabPage.Location = new System.Drawing.Point(4, 24);
            extractTabPage.Name = "extractTabPage";
            extractTabPage.Padding = new System.Windows.Forms.Padding(3);
            extractTabPage.Size = new System.Drawing.Size(1385, 456);
            extractTabPage.TabIndex = 1;
            extractTabPage.Text = "Extract";
            extractTabPage.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new System.Drawing.Point(1125, 20);
            label5.Name = "label5";
            label5.Size = new System.Drawing.Size(88, 15);
            label5.TabIndex = 10;
            label5.Text = "New Column";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new System.Drawing.Point(570, 20);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(53, 15);
            label3.TabIndex = 9;
            label3.Text = "Pattern";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label2.Location = new System.Drawing.Point(1, 20);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(60, 15);
            label2.TabIndex = 8;
            label2.Text = "Enabled";
            // 
            // toolTipBox
            // 
            toolTipBox.AutoPopDelay = 1000;
            toolTipBox.InitialDelay = 250;
            toolTipBox.ReshowDelay = 100;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new System.Drawing.Point(75, 20);
            label7.Name = "label7";
            label7.Size = new System.Drawing.Size(79, 15);
            label7.TabIndex = 12;
            label7.Text = "Rule Name";
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Location = new System.Drawing.Point(75, 20);
            label8.Name = "label8";
            label8.Size = new System.Drawing.Size(79, 15);
            label8.TabIndex = 13;
            label8.Text = "Rule Name";
            // 
            // DefineRulesForm
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            BackColor = System.Drawing.SystemColors.Window;
            ClientSize = new System.Drawing.Size(1414, 714);
            Controls.Add(selectedRowsGroupBox);
            Controls.Add(runButton);
            Controls.Add(clearButton);
            Controls.Add(discardButton);
            Controls.Add(saveButton);
            Controls.Add(sourceColumnGroupBox);
            Controls.Add(headline);
            Controls.Add(menuStrip1);
            Controls.Add(rulesTabControl);
            Icon = ((System.Drawing.Icon)(resources.GetObject("$Icon")));
            MainMenuStrip = menuStrip1;
            Name = "DefineRulesForm";
            Text = "Define Data Rules";
            sourceColumnGroupBox.ResumeLayout(false);
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            selectedRowsGroupBox.ResumeLayout(false);
            selectedRowsGroupBox.PerformLayout();
            extractRulesPanel.ResumeLayout(false);
            rulesTabControl.ResumeLayout(false);
            cleaningTabPage.ResumeLayout(false);
            cleaningTabPage.PerformLayout();
            cleaningRulesPanel.ResumeLayout(false);
            dateFormatTabPage.ResumeLayout(false);
            dateFormatTabPage.PerformLayout();
            groupBox1.ResumeLayout(false);
            extractTabPage.ResumeLayout(false);
            extractTabPage.PerformLayout();
            ResumeLayout(false);
            PerformLayout();

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
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
    }
}