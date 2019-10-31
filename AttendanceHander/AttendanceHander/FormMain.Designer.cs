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
            this.groupBoxMepStyle = new System.Windows.Forms.GroupBox();
            this.buttonTestMepStyle = new System.Windows.Forms.Button();
            this.label_StatusMepSty = new System.Windows.Forms.Label();
            this.buttonOpenMepTimesheet = new System.Windows.Forms.Button();
            this.groupBoxMultiTrans = new System.Windows.Forms.GroupBox();
            this.buttonTestMultiTrans = new System.Windows.Forms.Button();
            this.label_StatusMultiTrans = new System.Windows.Forms.Label();
            this.buttonOpenMultiTrans = new System.Windows.Forms.Button();
            this.groupBox_DailyTrans = new System.Windows.Forms.GroupBox();
            this.button_TestDailyTrans = new System.Windows.Forms.Button();
            this.label_StatusDailyTrans = new System.Windows.Forms.Label();
            this.button_OpenDailyTrans = new System.Windows.Forms.Button();
            this.groupBoxMepStyle.SuspendLayout();
            this.groupBoxMultiTrans.SuspendLayout();
            this.groupBox_DailyTrans.SuspendLayout();
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
            // groupBoxMepStyle
            // 
            this.groupBoxMepStyle.Controls.Add(this.buttonTestMepStyle);
            this.groupBoxMepStyle.Controls.Add(this.label_StatusMepSty);
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
            // label_StatusMepSty
            // 
            this.label_StatusMepSty.AutoSize = true;
            this.label_StatusMepSty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label_StatusMepSty.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_StatusMepSty.Location = new System.Drawing.Point(246, 65);
            this.label_StatusMepSty.Name = "label_StatusMepSty";
            this.label_StatusMepSty.Size = new System.Drawing.Size(325, 20);
            this.label_StatusMepSty.TabIndex = 8;
            this.label_StatusMepSty.Text = "Open Plumbers Time Sheet (Default MEP Style)";
            // 
            // buttonOpenMepTimesheet
            // 
            this.buttonOpenMepTimesheet.Location = new System.Drawing.Point(27, 55);
            this.buttonOpenMepTimesheet.Name = "buttonOpenMepTimesheet";
            this.buttonOpenMepTimesheet.Size = new System.Drawing.Size(143, 41);
            this.buttonOpenMepTimesheet.TabIndex = 7;
            this.buttonOpenMepTimesheet.Text = "Open and Process";
            this.buttonOpenMepTimesheet.UseVisualStyleBackColor = true;
            this.buttonOpenMepTimesheet.Click += new System.EventHandler(this.ButtonOpenMepTimesheet_Click);
            // 
            // groupBoxMultiTrans
            // 
            this.groupBoxMultiTrans.Controls.Add(this.buttonTestMultiTrans);
            this.groupBoxMultiTrans.Controls.Add(this.label_StatusMultiTrans);
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
            // label_StatusMultiTrans
            // 
            this.label_StatusMultiTrans.AutoSize = true;
            this.label_StatusMultiTrans.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label_StatusMultiTrans.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_StatusMultiTrans.Location = new System.Drawing.Point(246, 65);
            this.label_StatusMultiTrans.Name = "label_StatusMultiTrans";
            this.label_StatusMultiTrans.Size = new System.Drawing.Size(325, 20);
            this.label_StatusMultiTrans.TabIndex = 8;
            this.label_StatusMultiTrans.Text = "Open Plumbers Time Sheet (Default MEP Style)";
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
            // groupBox_DailyTrans
            // 
            this.groupBox_DailyTrans.Controls.Add(this.button_TestDailyTrans);
            this.groupBox_DailyTrans.Controls.Add(this.label_StatusDailyTrans);
            this.groupBox_DailyTrans.Controls.Add(this.button_OpenDailyTrans);
            this.groupBox_DailyTrans.Location = new System.Drawing.Point(12, 392);
            this.groupBox_DailyTrans.Name = "groupBox_DailyTrans";
            this.groupBox_DailyTrans.Size = new System.Drawing.Size(722, 124);
            this.groupBox_DailyTrans.TabIndex = 9;
            this.groupBox_DailyTrans.TabStop = false;
            this.groupBox_DailyTrans.Text = "DAILY TRANSACTIONS";
            // 
            // button_TestDailyTrans
            // 
            this.button_TestDailyTrans.Enabled = false;
            this.button_TestDailyTrans.Location = new System.Drawing.Point(615, 55);
            this.button_TestDailyTrans.Name = "button_TestDailyTrans";
            this.button_TestDailyTrans.Size = new System.Drawing.Size(101, 41);
            this.button_TestDailyTrans.TabIndex = 9;
            this.button_TestDailyTrans.Text = "Test MEP Style Timesheet";
            this.button_TestDailyTrans.UseVisualStyleBackColor = true;
            // 
            // label_StatusDailyTrans
            // 
            this.label_StatusDailyTrans.AutoSize = true;
            this.label_StatusDailyTrans.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label_StatusDailyTrans.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_StatusDailyTrans.Location = new System.Drawing.Point(246, 65);
            this.label_StatusDailyTrans.Name = "label_StatusDailyTrans";
            this.label_StatusDailyTrans.Size = new System.Drawing.Size(325, 20);
            this.label_StatusDailyTrans.TabIndex = 8;
            this.label_StatusDailyTrans.Text = "Open Plumbers Time Sheet (Default MEP Style)";
            // 
            // button_OpenDailyTrans
            // 
            this.button_OpenDailyTrans.Location = new System.Drawing.Point(27, 55);
            this.button_OpenDailyTrans.Name = "button_OpenDailyTrans";
            this.button_OpenDailyTrans.Size = new System.Drawing.Size(143, 41);
            this.button_OpenDailyTrans.TabIndex = 7;
            this.button_OpenDailyTrans.Text = "Open and Process";
            this.button_OpenDailyTrans.UseVisualStyleBackColor = true;
            this.button_OpenDailyTrans.Click += new System.EventHandler(this.Button_OpenDailyTrans_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(800, 542);
            this.Controls.Add(this.groupBox_DailyTrans);
            this.Controls.Add(this.groupBoxMultiTrans);
            this.Controls.Add(this.groupBoxMepStyle);
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
            this.groupBox_DailyTrans.ResumeLayout(false);
            this.groupBox_DailyTrans.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.GroupBox groupBoxMepStyle;
        private System.Windows.Forms.Button buttonTestMepStyle;
        private System.Windows.Forms.Label label_StatusMepSty;
        private System.Windows.Forms.Button buttonOpenMepTimesheet;
        private System.Windows.Forms.GroupBox groupBoxMultiTrans;
        private System.Windows.Forms.Button buttonTestMultiTrans;
        private System.Windows.Forms.Label label_StatusMultiTrans;
        private System.Windows.Forms.Button buttonOpenMultiTrans;
        private System.Windows.Forms.GroupBox groupBox_DailyTrans;
        private System.Windows.Forms.Button button_TestDailyTrans;
        private System.Windows.Forms.Label label_StatusDailyTrans;
        private System.Windows.Forms.Button button_OpenDailyTrans;
    }
}