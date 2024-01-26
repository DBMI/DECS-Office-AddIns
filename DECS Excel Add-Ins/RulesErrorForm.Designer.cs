namespace DECS_Excel_Add_Ins
{
    partial class RulesErrorForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RulesErrorForm));
            rulesErrorFormLabel = new System.Windows.Forms.Label();
            okButton = new System.Windows.Forms.Button();
            panel1 = new System.Windows.Forms.Panel();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // rulesErrorFormLabel
            // 
            rulesErrorFormLabel.AutoSize = true;
            rulesErrorFormLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            rulesErrorFormLabel.Location = new System.Drawing.Point(4, 9);
            rulesErrorFormLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            rulesErrorFormLabel.Name = "rulesErrorFormLabel";
            rulesErrorFormLabel.Size = new System.Drawing.Size(44, 16);
            rulesErrorFormLabel.TabIndex = 0;
            rulesErrorFormLabel.Text = "label1";
            // 
            // okButton
            // 
            okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            okButton.Location = new System.Drawing.Point(463, 304);
            okButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            okButton.Name = "okButton";
            okButton.Size = new System.Drawing.Size(100, 29);
            okButton.TabIndex = 1;
            okButton.Text = "OK";
            okButton.UseVisualStyleBackColor = true;
            okButton.Click += new System.EventHandler(okButton_Click);
            // 
            // panel1
            // 
            panel1.Controls.Add(rulesErrorFormLabel);
            panel1.Location = new System.Drawing.Point(12, 12);
            panel1.Name = "panel1";
            panel1.Size = new System.Drawing.Size(1004, 261);
            panel1.TabIndex = 2;
            // 
            // RulesErrorForm
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            BackColor = System.Drawing.Color.White;
            ClientSize = new System.Drawing.Size(1028, 359);
            Controls.Add(panel1);
            Controls.Add(okButton);
            Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            Icon = ((System.Drawing.Icon)(resources.GetObject("$Icon")));
            Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Name = "RulesErrorForm";
            Text = "Error in Rules";
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label rulesErrorFormLabel;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Panel panel1;
    }
}