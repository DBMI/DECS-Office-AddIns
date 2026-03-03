namespace DECS_Excel_Add_Ins
{
    partial class UseCalforniaOrAllUsaForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UseCalforniaOrAllUsaForm));
            this.label1 = new System.Windows.Forms.Label();
            this.radioButtonCalifornia = new System.Windows.Forms.RadioButton();
            this.radioButtonUSA = new System.Windows.Forms.RadioButton();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(147, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(270, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Which SVI data should we read?";
            // 
            // radioButtonCalifornia
            // 
            this.radioButtonCalifornia.AutoSize = true;
            this.radioButtonCalifornia.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonCalifornia.Location = new System.Drawing.Point(151, 113);
            this.radioButtonCalifornia.Name = "radioButtonCalifornia";
            this.radioButtonCalifornia.Size = new System.Drawing.Size(91, 20);
            this.radioButtonCalifornia.TabIndex = 1;
            this.radioButtonCalifornia.Text = "California";
            this.radioButtonCalifornia.UseVisualStyleBackColor = true;
            this.radioButtonCalifornia.CheckedChanged += new System.EventHandler(this.radioButtonCalifornia_CheckedChanged);
            // 
            // radioButtonUSA
            // 
            this.radioButtonUSA.AutoSize = true;
            this.radioButtonUSA.Checked = true;
            this.radioButtonUSA.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButtonUSA.Location = new System.Drawing.Point(151, 150);
            this.radioButtonUSA.Name = "radioButtonUSA";
            this.radioButtonUSA.Size = new System.Drawing.Size(56, 20);
            this.radioButtonUSA.TabIndex = 2;
            this.radioButtonUSA.TabStop = true;
            this.radioButtonUSA.Text = "USA";
            this.radioButtonUSA.UseVisualStyleBackColor = true;
            this.radioButtonUSA.CheckedChanged += new System.EventHandler(this.radioButtonUSA_CheckedChanged);
            // 
            // okButton
            // 
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.Location = new System.Drawing.Point(151, 222);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 3;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelButton.Location = new System.Drawing.Point(342, 222);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // UseCalforniaOrAllUsaForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(562, 313);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.radioButtonUSA);
            this.Controls.Add(this.radioButtonCalifornia);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "UseCalforniaOrAllUsaForm";
            this.Text = "UseCalforniaOrAllUsaForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton radioButtonCalifornia;
        private System.Windows.Forms.RadioButton radioButtonUSA;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
    }
}