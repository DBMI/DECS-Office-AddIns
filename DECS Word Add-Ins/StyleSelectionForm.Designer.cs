namespace DecsWordAddIns
{
    partial class StyleSelectionForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(StyleSelectionForm));
            this.logoPictureBox = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.icdStyleGroupBox = new System.Windows.Forms.GroupBox();
            this.listRadioButton = new System.Windows.Forms.RadioButton();
            this.caseRadioButton = new System.Windows.Forms.RadioButton();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.icdStyleGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // logoPictureBox
            // 
            this.logoPictureBox.Image = global::DecsWordAddIns.Properties.Resources.school_of_medicine;
            this.logoPictureBox.InitialImage = global::DecsWordAddIns.Properties.Resources.school_of_medicine;
            this.logoPictureBox.Location = new System.Drawing.Point(1, 3);
            this.logoPictureBox.Name = "logoPictureBox";
            this.logoPictureBox.Size = new System.Drawing.Size(256, 88);
            this.logoPictureBox.TabIndex = 7;
            this.logoPictureBox.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(492, 66);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(183, 25);
            this.label1.TabIndex = 8;
            this.label1.Text = "Select ICD Style";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::DecsWordAddIns.Properties.Resources.ICD_case_statements;
            this.pictureBox1.Location = new System.Drawing.Point(41, 99);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(494, 227);
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::DecsWordAddIns.Properties.Resources.ICD_list_statement;
            this.pictureBox2.Location = new System.Drawing.Point(600, 99);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(511, 227);
            this.pictureBox2.TabIndex = 10;
            this.pictureBox2.TabStop = false;
            // 
            // icdStyleGroupBox
            // 
            this.icdStyleGroupBox.Controls.Add(this.listRadioButton);
            this.icdStyleGroupBox.Controls.Add(this.pictureBox2);
            this.icdStyleGroupBox.Controls.Add(this.caseRadioButton);
            this.icdStyleGroupBox.Controls.Add(this.pictureBox1);
            this.icdStyleGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.icdStyleGroupBox.Location = new System.Drawing.Point(12, 115);
            this.icdStyleGroupBox.Name = "icdStyleGroupBox";
            this.icdStyleGroupBox.Size = new System.Drawing.Size(1142, 352);
            this.icdStyleGroupBox.TabIndex = 11;
            this.icdStyleGroupBox.TabStop = false;
            this.icdStyleGroupBox.Text = "ICD style";
            // 
            // listRadioButton
            // 
            this.listRadioButton.AutoSize = true;
            this.listRadioButton.Location = new System.Drawing.Point(600, 42);
            this.listRadioButton.Name = "listRadioButton";
            this.listRadioButton.Size = new System.Drawing.Size(78, 20);
            this.listRadioButton.TabIndex = 1;
            this.listRadioButton.Text = "ICD List";
            this.listRadioButton.UseVisualStyleBackColor = true;
            // 
            // caseRadioButton
            // 
            this.caseRadioButton.AutoSize = true;
            this.caseRadioButton.Checked = true;
            this.caseRadioButton.Location = new System.Drawing.Point(41, 40);
            this.caseRadioButton.Name = "caseRadioButton";
            this.caseRadioButton.Size = new System.Drawing.Size(217, 20);
            this.caseRadioButton.TabIndex = 0;
            this.caseRadioButton.TabStop = true;
            this.caseRadioButton.Text = "Individual CASE Statements";
            this.caseRadioButton.UseVisualStyleBackColor = true;
            this.caseRadioButton.CheckedChanged += new System.EventHandler(this.StyleRadioButton_CheckedChanged);
            // 
            // okButton
            // 
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.Location = new System.Drawing.Point(447, 500);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(100, 40);
            this.okButton.TabIndex = 12;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelButton.Location = new System.Drawing.Point(612, 500);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(100, 40);
            this.cancelButton.TabIndex = 13;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // StyleSelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1185, 569);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.icdStyleGroupBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.logoPictureBox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "StyleSelectionForm";
            this.Text = "Style Selection Form";
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.icdStyleGroupBox.ResumeLayout(false);
            this.icdStyleGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox logoPictureBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.GroupBox icdStyleGroupBox;
        private System.Windows.Forms.RadioButton listRadioButton;
        private System.Windows.Forms.RadioButton caseRadioButton;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
    }
}