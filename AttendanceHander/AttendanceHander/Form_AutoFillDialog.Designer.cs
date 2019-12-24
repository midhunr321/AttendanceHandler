namespace AttendanceHander
{
    partial class Form_AutoFillDialog
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
            this.label1 = new System.Windows.Forms.Label();
            this.checkBox_autoFillSite = new System.Windows.Forms.CheckBox();
            this.button_proceed = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.checkBox_autoFillSite);
            this.groupBox1.Location = new System.Drawing.Point(29, 21);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(398, 190);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(32, 77);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(338, 45);
            this.label1.TabIndex = 1;
            this.label1.Text = "Software will automatically understand the majority site number of a particular e" +
    "mployee and place it on fridays and holidays";
            // 
            // checkBox_autoFillSite
            // 
            this.checkBox_autoFillSite.AutoSize = true;
            this.checkBox_autoFillSite.Location = new System.Drawing.Point(32, 41);
            this.checkBox_autoFillSite.Name = "checkBox_autoFillSite";
            this.checkBox_autoFillSite.Size = new System.Drawing.Size(338, 17);
            this.checkBox_autoFillSite.TabIndex = 0;
            this.checkBox_autoFillSite.Text = "Auto-Fill Site No.s in Multiple Transactions for Holidays and Fridays";
            this.checkBox_autoFillSite.UseVisualStyleBackColor = true;
            // 
            // button_proceed
            // 
            this.button_proceed.Location = new System.Drawing.Point(164, 230);
            this.button_proceed.Name = "button_proceed";
            this.button_proceed.Size = new System.Drawing.Size(144, 36);
            this.button_proceed.TabIndex = 1;
            this.button_proceed.Text = "Proceed";
            this.button_proceed.UseVisualStyleBackColor = true;
            this.button_proceed.Click += new System.EventHandler(this.Button_proceed_Click);
            // 
            // Form_AutoFillDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(466, 282);
            this.Controls.Add(this.button_proceed);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form_AutoFillDialog";
            this.Text = "Auto Fill Site No. Dialog";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBox_autoFillSite;
        private System.Windows.Forms.Button button_proceed;
    }
}