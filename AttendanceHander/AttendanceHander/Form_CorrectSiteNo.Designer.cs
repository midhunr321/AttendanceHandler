namespace AttendanceHander
{
    partial class Form_CorrectSiteNo
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButton_fullName = new System.Windows.Forms.RadioButton();
            this.radioButton_shortName = new System.Windows.Forms.RadioButton();
            this.button_ok = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button_ok);
            this.groupBox1.Controls.Add(this.radioButton_shortName);
            this.groupBox1.Controls.Add(this.radioButton_fullName);
            this.groupBox1.Location = new System.Drawing.Point(12, 21);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(333, 159);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Site No Style";
            // 
            // radioButton_fullName
            // 
            this.radioButton_fullName.AutoSize = true;
            this.radioButton_fullName.Location = new System.Drawing.Point(96, 45);
            this.radioButton_fullName.Name = "radioButton_fullName";
            this.radioButton_fullName.Size = new System.Drawing.Size(181, 17);
            this.radioButton_fullName.TabIndex = 0;
            this.radioButton_fullName.TabStop = true;
            this.radioButton_fullName.Text = "Full Site No; Example = M269-D2";
            this.radioButton_fullName.UseVisualStyleBackColor = true;
            // 
            // radioButton_shortName
            // 
            this.radioButton_shortName.AutoSize = true;
            this.radioButton_shortName.Location = new System.Drawing.Point(96, 77);
            this.radioButton_shortName.Name = "radioButton_shortName";
            this.radioButton_shortName.Size = new System.Drawing.Size(173, 17);
            this.radioButton_shortName.TabIndex = 1;
            this.radioButton_shortName.TabStop = true;
            this.radioButton_shortName.Text = "Short Site No; Example = M269";
            this.radioButton_shortName.UseVisualStyleBackColor = true;
            // 
            // button_ok
            // 
            this.button_ok.Location = new System.Drawing.Point(96, 121);
            this.button_ok.Name = "button_ok";
            this.button_ok.Size = new System.Drawing.Size(142, 32);
            this.button_ok.TabIndex = 2;
            this.button_ok.Text = "OK";
            this.button_ok.UseVisualStyleBackColor = true;
            this.button_ok.Click += new System.EventHandler(this.Button_ok_Click);
            // 
            // Form_CorrectSiteNo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.ClientSize = new System.Drawing.Size(367, 192);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form_CorrectSiteNo";
            this.Text = "Form_CorrectSiteNo";
            this.TopMost = true;
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button_ok;
        private System.Windows.Forms.RadioButton radioButton_shortName;
        private System.Windows.Forms.RadioButton radioButton_fullName;
    }
}