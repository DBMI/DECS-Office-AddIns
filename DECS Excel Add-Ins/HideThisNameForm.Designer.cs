namespace DECS_Excel_Add_Ins
{
    partial class HideThisNameForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HideThisNameForm));
            this.label1 = new System.Windows.Forms.Label();
            this.yesButton = new System.Windows.Forms.Button();
            this.noButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.similarNamesListBox = new System.Windows.Forms.ListBox();
            this.linkNameButton = new System.Windows.Forms.Button();
            this.contextRichTextBox = new System.Windows.Forms.RichTextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(149, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(295, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Should we hide this name?";
            // 
            // yesButton
            // 
            this.yesButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.yesButton.Location = new System.Drawing.Point(137, 146);
            this.yesButton.Name = "yesButton";
            this.yesButton.Size = new System.Drawing.Size(75, 23);
            this.yesButton.TabIndex = 2;
            this.yesButton.Text = "Hide";
            this.yesButton.UseVisualStyleBackColor = true;
            this.yesButton.Click += new System.EventHandler(this.NewButton_Click);
            // 
            // noButton
            // 
            this.noButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.noButton.Location = new System.Drawing.Point(137, 175);
            this.noButton.Name = "noButton";
            this.noButton.Size = new System.Drawing.Size(75, 23);
            this.noButton.TabIndex = 3;
            this.noButton.Text = "Ignore";
            this.noButton.UseVisualStyleBackColor = true;
            this.noButton.Click += new System.EventHandler(this.IgnoreButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelButton.Location = new System.Drawing.Point(137, 204);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // similarNamesListBox
            // 
            this.similarNamesListBox.FormattingEnabled = true;
            this.similarNamesListBox.Location = new System.Drawing.Point(222, 117);
            this.similarNamesListBox.Name = "similarNamesListBox";
            this.similarNamesListBox.Size = new System.Drawing.Size(209, 108);
            this.similarNamesListBox.TabIndex = 7;
            this.similarNamesListBox.Click += new System.EventHandler(this.similarNamesListBox_Click);
            // 
            // linkNameButton
            // 
            this.linkNameButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkNameButton.Location = new System.Drawing.Point(137, 117);
            this.linkNameButton.Name = "linkNameButton";
            this.linkNameButton.Size = new System.Drawing.Size(75, 23);
            this.linkNameButton.TabIndex = 8;
            this.linkNameButton.Text = "Link";
            this.linkNameButton.UseVisualStyleBackColor = true;
            this.linkNameButton.Click += new System.EventHandler(this.LinkButton_Click);
            // 
            // contextRichTextBox
            // 
            this.contextRichTextBox.AutoWordSelection = true;
            this.contextRichTextBox.BackColor = System.Drawing.Color.White;
            this.contextRichTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.contextRichTextBox.Location = new System.Drawing.Point(32, 63);
            this.contextRichTextBox.Name = "contextRichTextBox";
            this.contextRichTextBox.ReadOnly = true;
            this.contextRichTextBox.Size = new System.Drawing.Size(528, 29);
            this.contextRichTextBox.TabIndex = 9;
            this.contextRichTextBox.Text = "";
            this.contextRichTextBox.ZoomFactor = 2F;
            this.contextRichTextBox.SelectionChanged += new System.EventHandler(this.UserSelectedPartOfProvidedText);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(439, 159);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 10;
            this.button1.Text = "Show All";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.ShowAllButton_Click);
            // 
            // HideThisNameForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(605, 296);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.contextRichTextBox);
            this.Controls.Add(this.linkNameButton);
            this.Controls.Add(this.similarNamesListBox);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.noButton);
            this.Controls.Add(this.yesButton);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "HideThisNameForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Hide This Name?";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button yesButton;
        private System.Windows.Forms.Button noButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.ListBox similarNamesListBox;
        private System.Windows.Forms.Button linkNameButton;
        private System.Windows.Forms.RichTextBox contextRichTextBox;
        private System.Windows.Forms.Button button1;
    }
}