namespace DecsWordAddIns
{
    partial class DeliveryTypeForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.oneDriveRadioButton = new System.Windows.Forms.RadioButton();
            this.vrdRadioButton = new System.Windows.Forms.RadioButton();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.logoPictureBox = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(119, 128);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(275, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "How will the results be delivered?";
            // 
            // oneDriveRadioButton
            // 
            this.oneDriveRadioButton.AutoSize = true;
            this.oneDriveRadioButton.Checked = true;
            this.oneDriveRadioButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.oneDriveRadioButton.Location = new System.Drawing.Point(165, 168);
            this.oneDriveRadioButton.Name = "oneDriveRadioButton";
            this.oneDriveRadioButton.Size = new System.Drawing.Size(85, 20);
            this.oneDriveRadioButton.TabIndex = 1;
            this.oneDriveRadioButton.TabStop = true;
            this.oneDriveRadioButton.Text = "One Drive";
            this.oneDriveRadioButton.UseVisualStyleBackColor = true;
            this.oneDriveRadioButton.CheckedChanged += new System.EventHandler(this.oneDriveRadioButton_CheckedChanged);
            // 
            // vrdRadioButton
            // 
            this.vrdRadioButton.AutoSize = true;
            this.vrdRadioButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.vrdRadioButton.Location = new System.Drawing.Point(165, 195);
            this.vrdRadioButton.Name = "vrdRadioButton";
            this.vrdRadioButton.Size = new System.Drawing.Size(218, 20);
            this.vrdRadioButton.TabIndex = 2;
            this.vrdRadioButton.Text = "Virtual Research Desktop (VRD)";
            this.vrdRadioButton.UseVisualStyleBackColor = true;
            this.vrdRadioButton.CheckedChanged += new System.EventHandler(this.vrdRadioButton_CheckedChanged);
            // 
            // okButton
            // 
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.Location = new System.Drawing.Point(144, 288);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(92, 34);
            this.okButton.TabIndex = 3;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancelButton.Location = new System.Drawing.Point(302, 288);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(92, 34);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // logoPictureBox
            // 
            this.logoPictureBox.Image = global::DecsWordAddIns.Properties.Resources.school_of_medicine;
            this.logoPictureBox.InitialImage = global::DecsWordAddIns.Properties.Resources.school_of_medicine;
            this.logoPictureBox.Location = new System.Drawing.Point(-6, -3);
            this.logoPictureBox.Name = "logoPictureBox";
            this.logoPictureBox.Size = new System.Drawing.Size(256, 88);
            this.logoPictureBox.TabIndex = 5;
            this.logoPictureBox.TabStop = false;
            // 
            // DeliveryTypeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(545, 398);
            this.Controls.Add(this.logoPictureBox);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.vrdRadioButton);
            this.Controls.Add(this.oneDriveRadioButton);
            this.Controls.Add(this.label1);
            this.Name = "DeliveryTypeForm";
            this.Text = "Delivery Type";
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton oneDriveRadioButton;
        private System.Windows.Forms.RadioButton vrdRadioButton;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.PictureBox logoPictureBox;
    }
}