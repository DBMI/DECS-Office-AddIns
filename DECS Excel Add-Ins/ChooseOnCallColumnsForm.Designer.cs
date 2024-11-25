namespace DECS_Excel_Add_Ins
{
    partial class ChooseOnCallColumnsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChooseOnCallColumnsForm));
            this.label1 = new System.Windows.Forms.Label();
            this.onCallDateColumnListBox = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.onCallNameColumnListBox = new System.Windows.Forms.ListBox();
            this.onCallQuitButton = new System.Windows.Forms.Button();
            this.onCallRunButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(227, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(320, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Choose Date, Name Columns";
            // 
            // onCallDateColumnListBox
            // 
            this.onCallDateColumnListBox.FormattingEnabled = true;
            this.onCallDateColumnListBox.Location = new System.Drawing.Point(134, 131);
            this.onCallDateColumnListBox.Name = "onCallDateColumnListBox";
            this.onCallDateColumnListBox.Size = new System.Drawing.Size(120, 134);
            this.onCallDateColumnListBox.TabIndex = 2;
            this.onCallDateColumnListBox.SelectedIndexChanged += new System.EventHandler(this.DateColumn_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(130, 108);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 20);
            this.label2.TabIndex = 3;
            this.label2.Text = "Date column";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(554, 108);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(117, 20);
            this.label3.TabIndex = 4;
            this.label3.Text = "Name column";
            // 
            // onCallNameColumnListBox
            // 
            this.onCallNameColumnListBox.FormattingEnabled = true;
            this.onCallNameColumnListBox.Location = new System.Drawing.Point(558, 131);
            this.onCallNameColumnListBox.Name = "onCallNameColumnListBox";
            this.onCallNameColumnListBox.Size = new System.Drawing.Size(120, 134);
            this.onCallNameColumnListBox.TabIndex = 5;
            this.onCallNameColumnListBox.SelectedIndexChanged += new System.EventHandler(this.NameColumn_SelectedIndexChanged);
            // 
            // onCallQuitButton
            // 
            this.onCallQuitButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.onCallQuitButton.Location = new System.Drawing.Point(444, 316);
            this.onCallQuitButton.Name = "onCallQuitButton";
            this.onCallQuitButton.Size = new System.Drawing.Size(120, 40);
            this.onCallQuitButton.TabIndex = 7;
            this.onCallQuitButton.Text = "Quit";
            this.onCallQuitButton.UseVisualStyleBackColor = true;
            this.onCallQuitButton.Click += new System.EventHandler(this.QuitButton_Click);
            // 
            // onCallRunButton
            // 
            this.onCallRunButton.Enabled = false;
            this.onCallRunButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.onCallRunButton.Location = new System.Drawing.Point(238, 316);
            this.onCallRunButton.Name = "onCallRunButton";
            this.onCallRunButton.Size = new System.Drawing.Size(120, 40);
            this.onCallRunButton.TabIndex = 6;
            this.onCallRunButton.Text = "Run";
            this.onCallRunButton.UseVisualStyleBackColor = true;
            this.onCallRunButton.Click += new System.EventHandler(this.RunButton_Click);
            // 
            // ChooseOnCallColumnsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(800, 403);
            this.Controls.Add(this.onCallQuitButton);
            this.Controls.Add(this.onCallRunButton);
            this.Controls.Add(this.onCallNameColumnListBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.onCallDateColumnListBox);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ChooseOnCallColumnsForm";
            this.Text = "Choose On-Call Columns Form";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox onCallDateColumnListBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox onCallNameColumnListBox;
        private System.Windows.Forms.Button onCallQuitButton;
        private System.Windows.Forms.Button onCallRunButton;
    }
}