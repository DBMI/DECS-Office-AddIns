namespace DECS_Excel_Add_Ins
{
    partial class MatchTextForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MatchTextForm));
            this.label1 = new System.Windows.Forms.Label();
            this.redcapMessageColumnsListBox = new System.Windows.Forms.ListBox();
            this.artMessageColumnsListBox = new System.Windows.Forms.ListBox();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.redcapSheetsListBox = new System.Windows.Forms.ListBox();
            this.artSheetsListBox = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.artIdColumnsListBox = new System.Windows.Forms.ListBox();
            this.label7 = new System.Windows.Forms.Label();
            this.redcapIdColumnsListBox = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(372, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(227, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Match Text Columns";
            // 
            // redcapMessageColumnsListBox
            // 
            this.redcapMessageColumnsListBox.Enabled = false;
            this.redcapMessageColumnsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.redcapMessageColumnsListBox.FormattingEnabled = true;
            this.redcapMessageColumnsListBox.ItemHeight = 16;
            this.redcapMessageColumnsListBox.Location = new System.Drawing.Point(250, 210);
            this.redcapMessageColumnsListBox.Name = "redcapMessageColumnsListBox";
            this.redcapMessageColumnsListBox.Size = new System.Drawing.Size(200, 196);
            this.redcapMessageColumnsListBox.TabIndex = 1;
            this.redcapMessageColumnsListBox.SelectedIndexChanged += new System.EventHandler(this.EnableWhenReady);
            // 
            // artMessageColumnsListBox
            // 
            this.artMessageColumnsListBox.Enabled = false;
            this.artMessageColumnsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.artMessageColumnsListBox.FormattingEnabled = true;
            this.artMessageColumnsListBox.ItemHeight = 16;
            this.artMessageColumnsListBox.Location = new System.Drawing.Point(744, 210);
            this.artMessageColumnsListBox.Name = "artMessageColumnsListBox";
            this.artMessageColumnsListBox.Size = new System.Drawing.Size(200, 196);
            this.artMessageColumnsListBox.TabIndex = 2;
            this.artMessageColumnsListBox.SelectedIndexChanged += new System.EventHandler(this.EnableWhenReady);
            // 
            // okButton
            // 
            this.okButton.Enabled = false;
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.Location = new System.Drawing.Point(350, 445);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(100, 40);
            this.okButton.TabIndex = 3;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelButton.Location = new System.Drawing.Point(530, 445);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(100, 40);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // redcapSheetsListBox
            // 
            this.redcapSheetsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.redcapSheetsListBox.FormattingEnabled = true;
            this.redcapSheetsListBox.ItemHeight = 16;
            this.redcapSheetsListBox.Location = new System.Drawing.Point(113, 102);
            this.redcapSheetsListBox.Name = "redcapSheetsListBox";
            this.redcapSheetsListBox.Size = new System.Drawing.Size(200, 68);
            this.redcapSheetsListBox.TabIndex = 5;
            this.redcapSheetsListBox.SelectedIndexChanged += new System.EventHandler(this.redcapSheetsListBox_SelectedIndexChanged);
            // 
            // artSheetsListBox
            // 
            this.artSheetsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.artSheetsListBox.FormattingEnabled = true;
            this.artSheetsListBox.ItemHeight = 16;
            this.artSheetsListBox.Location = new System.Drawing.Point(648, 102);
            this.artSheetsListBox.Name = "artSheetsListBox";
            this.artSheetsListBox.Size = new System.Drawing.Size(200, 68);
            this.artSheetsListBox.TabIndex = 6;
            this.artSheetsListBox.SelectedIndexChanged += new System.EventHandler(this.artSheetsListBox_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(157, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(111, 16);
            this.label2.TabIndex = 7;
            this.label2.Text = "REDCap Sheet";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(684, 74);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(136, 16);
            this.label3.TabIndex = 8;
            this.label3.Text = "ART SPARK Sheet";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(255, 186);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(190, 16);
            this.label4.TabIndex = 9;
            this.label4.Text = "REDCap Message Column";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(736, 186);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(215, 16);
            this.label5.TabIndex = 10;
            this.label5.Text = "ART SPARK Message Column";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(549, 186);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(166, 16);
            this.label6.TabIndex = 12;
            this.label6.Text = "ART SPARK ID Column";
            // 
            // artIdColumnsListBox
            // 
            this.artIdColumnsListBox.Enabled = false;
            this.artIdColumnsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.artIdColumnsListBox.FormattingEnabled = true;
            this.artIdColumnsListBox.ItemHeight = 16;
            this.artIdColumnsListBox.Location = new System.Drawing.Point(530, 210);
            this.artIdColumnsListBox.Name = "artIdColumnsListBox";
            this.artIdColumnsListBox.Size = new System.Drawing.Size(200, 196);
            this.artIdColumnsListBox.TabIndex = 11;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(38, 186);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(141, 16);
            this.label7.TabIndex = 14;
            this.label7.Text = "REDCap ID Column";
            // 
            // redcapIdColumnsListBox
            // 
            this.redcapIdColumnsListBox.Enabled = false;
            this.redcapIdColumnsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.redcapIdColumnsListBox.FormattingEnabled = true;
            this.redcapIdColumnsListBox.ItemHeight = 16;
            this.redcapIdColumnsListBox.Location = new System.Drawing.Point(33, 210);
            this.redcapIdColumnsListBox.Name = "redcapIdColumnsListBox";
            this.redcapIdColumnsListBox.Size = new System.Drawing.Size(200, 196);
            this.redcapIdColumnsListBox.TabIndex = 13;
            // 
            // MatchTextForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(986, 539);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.redcapIdColumnsListBox);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.artIdColumnsListBox);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.artSheetsListBox);
            this.Controls.Add(this.redcapSheetsListBox);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.artMessageColumnsListBox);
            this.Controls.Add(this.redcapMessageColumnsListBox);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MatchTextForm";
            this.Text = "MatchTextForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox redcapMessageColumnsListBox;
        private System.Windows.Forms.ListBox artMessageColumnsListBox;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.ListBox redcapSheetsListBox;
        private System.Windows.Forms.ListBox artSheetsListBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ListBox artIdColumnsListBox;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ListBox redcapIdColumnsListBox;
    }
}