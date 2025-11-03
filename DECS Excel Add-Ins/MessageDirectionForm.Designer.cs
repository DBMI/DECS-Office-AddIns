namespace DECS_Excel_Add_Ins
{
    partial class MessageDirectionForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MessageDirectionForm));
            this.label1 = new System.Windows.Forms.Label();
            this.fromPatientRadioButton = new System.Windows.Forms.RadioButton();
            this.toPatientRadioButton = new System.Windows.Forms.RadioButton();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.noneRadioButton = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(24, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(295, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Choose Message Direction";
            // 
            // fromPatientRadioButton
            // 
            this.fromPatientRadioButton.AutoSize = true;
            this.fromPatientRadioButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fromPatientRadioButton.Location = new System.Drawing.Point(109, 87);
            this.fromPatientRadioButton.Name = "fromPatientRadioButton";
            this.fromPatientRadioButton.Size = new System.Drawing.Size(112, 20);
            this.fromPatientRadioButton.TabIndex = 1;
            this.fromPatientRadioButton.Text = "From Patient";
            this.fromPatientRadioButton.UseVisualStyleBackColor = true;
            this.fromPatientRadioButton.Click += new System.EventHandler(this.fromPatientRadioButton_Clicked);
            // 
            // toPatientRadioButton
            // 
            this.toPatientRadioButton.AutoSize = true;
            this.toPatientRadioButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toPatientRadioButton.Location = new System.Drawing.Point(109, 130);
            this.toPatientRadioButton.Name = "toPatientRadioButton";
            this.toPatientRadioButton.Size = new System.Drawing.Size(96, 20);
            this.toPatientRadioButton.TabIndex = 2;
            this.toPatientRadioButton.Text = "To Patient";
            this.toPatientRadioButton.UseVisualStyleBackColor = true;
            this.toPatientRadioButton.Click += new System.EventHandler(this.toPatientRadioButton_Click);
            // 
            // okButton
            // 
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.Location = new System.Drawing.Point(73, 225);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 28);
            this.okButton.TabIndex = 3;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelButton.Location = new System.Drawing.Point(183, 225);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 28);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // noneRadioButton
            // 
            this.noneRadioButton.AutoSize = true;
            this.noneRadioButton.Checked = true;
            this.noneRadioButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.noneRadioButton.Location = new System.Drawing.Point(109, 169);
            this.noneRadioButton.Name = "noneRadioButton";
            this.noneRadioButton.Size = new System.Drawing.Size(106, 20);
            this.noneRadioButton.TabIndex = 5;
            this.noneRadioButton.TabStop = true;
            this.noneRadioButton.Text = "I don\'t know";
            this.noneRadioButton.UseVisualStyleBackColor = true;
            this.noneRadioButton.Click += new System.EventHandler(this.noneRadioButton_Click);
            // 
            // MessageDirectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(346, 275);
            this.Controls.Add(this.noneRadioButton);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.toPatientRadioButton);
            this.Controls.Add(this.fromPatientRadioButton);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MessageDirectionForm";
            this.Text = "MessageDirectionForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton fromPatientRadioButton;
        private System.Windows.Forms.RadioButton toPatientRadioButton;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.RadioButton noneRadioButton;
    }
}