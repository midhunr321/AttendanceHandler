namespace AttendanceHander
{
    partial class FormMain
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
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.groupBoxMepStyle = new System.Windows.Forms.GroupBox();
            this.buttonTestMepStyle = new System.Windows.Forms.Button();
            this.labelMepStyle = new System.Windows.Forms.Label();
            this.buttonOpenMepTimesheet = new System.Windows.Forms.Button();
            this.groupBoxMultiTrans = new System.Windows.Forms.GroupBox();
            this.buttonTestMultiTrans = new System.Windows.Forms.Button();
            this.labelMultiTrans = new System.Windows.Forms.Label();
            this.buttonOpenMultiTrans = new System.Windows.Forms.Button();
            this.groupBoxMepStyle.SuspendLayout();
            this.groupBoxMultiTrans.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(52, 33);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(192, 73);
            this.button1.TabIndex = 0;
            this.button1.Text = "Add site Nos to current";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(228, 402);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 18);
            this.label2.TabIndex = 4;
            this.label2.Text = "label2";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(12, 379);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(143, 41);
            this.button3.TabIndex = 3;
            this.button3.Text = "Open";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // groupBoxMepStyle
            // 
            this.groupBoxMepStyle.Controls.Add(this.buttonTestMepStyle);
            this.groupBoxMepStyle.Controls.Add(this.labelMepStyle);
            this.groupBoxMepStyle.Controls.Add(this.buttonOpenMepTimesheet);
            this.groupBoxMepStyle.Location = new System.Drawing.Point(12, 121);
            this.groupBoxMepStyle.Name = "groupBoxMepStyle";
            this.groupBoxMepStyle.Size = new System.Drawing.Size(722, 124);
            this.groupBoxMepStyle.TabIndex = 7;
            this.groupBoxMepStyle.TabStop = false;
            this.groupBoxMepStyle.Text = "MEP STYLE TIME SHEET";
            this.groupBoxMepStyle.Enter += new System.EventHandler(this.GroupBox1_Enter);
            // 
            // buttonTestMepStyle
            // 
            this.buttonTestMepStyle.Enabled = false;
            this.buttonTestMepStyle.Location = new System.Drawing.Point(615, 55);
            this.buttonTestMepStyle.Name = "buttonTestMepStyle";
            this.buttonTestMepStyle.Size = new System.Drawing.Size(101, 41);
            this.buttonTestMepStyle.TabIndex = 9;
            this.buttonTestMepStyle.Text = "Test MEP Style Timesheet";
            this.buttonTestMepStyle.UseVisualStyleBackColor = true;
            // 
            // labelMepStyle
            // 
            this.labelMepStyle.AutoSize = true;
            this.labelMepStyle.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelMepStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelMepStyle.Location = new System.Drawing.Point(246, 65);
            this.labelMepStyle.Name = "labelMepStyle";
            this.labelMepStyle.Size = new System.Drawing.Size(325, 20);
            this.labelMepStyle.TabIndex = 8;
            this.labelMepStyle.Text = "Open Plumbers Time Sheet (Default MEP Style)";
            // 
            // buttonOpenMepTimesheet
            // 
            this.buttonOpenMepTimesheet.Location = new System.Drawing.Point(27, 55);
            this.buttonOpenMepTimesheet.Name = "buttonOpenMepTimesheet";
            this.buttonOpenMepTimesheet.Size = new System.Drawing.Size(143, 41);
            this.buttonOpenMepTimesheet.TabIndex = 7;
            this.buttonOpenMepTimesheet.Text = "Open and Process";
            this.buttonOpenMepTimesheet.UseVisualStyleBackColor = true;
            // 
            // groupBoxMultiTrans
            // 
            this.groupBoxMultiTrans.Controls.Add(this.buttonTestMultiTrans);
            this.groupBoxMultiTrans.Controls.Add(this.labelMultiTrans);
            this.groupBoxMultiTrans.Controls.Add(this.buttonOpenMultiTrans);
            this.groupBoxMultiTrans.Location = new System.Drawing.Point(12, 251);
            this.groupBoxMultiTrans.Name = "groupBoxMultiTrans";
            this.groupBoxMultiTrans.Size = new System.Drawing.Size(722, 124);
            this.groupBoxMultiTrans.TabIndex = 8;
            this.groupBoxMultiTrans.TabStop = false;
            this.groupBoxMultiTrans.Text = "MULTIPLE TRANSACTION";
            // 
            // buttonTestMultiTrans
            // 
            this.buttonTestMultiTrans.Enabled = false;
            this.buttonTestMultiTrans.Location = new System.Drawing.Point(615, 55);
            this.buttonTestMultiTrans.Name = "buttonTestMultiTrans";
            this.buttonTestMultiTrans.Size = new System.Drawing.Size(101, 41);
            this.buttonTestMultiTrans.TabIndex = 9;
            this.buttonTestMultiTrans.Text = "Test MEP Style Timesheet";
            this.buttonTestMultiTrans.UseVisualStyleBackColor = true;
            // 
            // labelMultiTrans
            // 
            this.labelMultiTrans.AutoSize = true;
            this.labelMultiTrans.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelMultiTrans.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelMultiTrans.Location = new System.Drawing.Point(246, 65);
            this.labelMultiTrans.Name = "labelMultiTrans";
            this.labelMultiTrans.Size = new System.Drawing.Size(325, 20);
            this.labelMultiTrans.TabIndex = 8;
            this.labelMultiTrans.Text = "Open Plumbers Time Sheet (Default MEP Style)";
            // 
            // buttonOpenMultiTrans
            // 
            this.buttonOpenMultiTrans.Location = new System.Drawing.Point(27, 55);
            this.buttonOpenMultiTrans.Name = "buttonOpenMultiTrans";
            this.buttonOpenMultiTrans.Size = new System.Drawing.Size(143, 41);
            this.buttonOpenMultiTrans.TabIndex = 7;
            this.buttonOpenMultiTrans.Text = "Open and Process";
            this.buttonOpenMultiTrans.UseVisualStyleBackColor = true;
            this.buttonOpenMultiTrans.Click += new System.EventHandler(this.buttonOpenMultiTrans_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(800, 514);
            this.Controls.Add(this.groupBoxMultiTrans);
            this.Controls.Add(this.groupBoxMepStyle);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Name = "FormMain";
            this.Text = "Attendance Handler";
            this.Activated += new System.EventHandler(this.FormMain_Activated);
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.groupBoxMepStyle.ResumeLayout(false);
            this.groupBoxMepStyle.PerformLayout();
            this.groupBoxMultiTrans.ResumeLayout(false);
            this.groupBoxMultiTrans.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.GroupBox groupBoxMepStyle;
        private System.Windows.Forms.Button buttonTestMepStyle;
        private System.Windows.Forms.Label labelMepStyle;
        private System.Windows.Forms.Button buttonOpenMepTimesheet;
        private System.Windows.Forms.GroupBox groupBoxMultiTrans;
        private System.Windows.Forms.Button buttonTestMultiTrans;
        private System.Windows.Forms.Label labelMultiTrans;
        private System.Windows.Forms.Button buttonOpenMultiTrans;
    }
}