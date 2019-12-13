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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBoxMepStyle = new System.Windows.Forms.GroupBox();
            this.button_clearMepStyle = new System.Windows.Forms.Button();
            this.buttonTestMepStyle = new System.Windows.Forms.Button();
            this.label_StatusMepSty = new System.Windows.Forms.Label();
            this.buttonOpenMepTimesheet = new System.Windows.Forms.Button();
            this.groupBoxMultiTrans = new System.Windows.Forms.GroupBox();
            this.button_clearMultiTrans = new System.Windows.Forms.Button();
            this.buttonTestMultiTrans = new System.Windows.Forms.Button();
            this.label_StatusMultiTrans = new System.Windows.Forms.Label();
            this.buttonOpenMultiTrans = new System.Windows.Forms.Button();
            this.groupBox_DailyTrans = new System.Windows.Forms.GroupBox();
            this.button_clearDailyTrans = new System.Windows.Forms.Button();
            this.button_TestDailyTrans = new System.Windows.Forms.Button();
            this.label_StatusDailyTrans = new System.Windows.Forms.Label();
            this.button_OpenDailyTrans = new System.Windows.Forms.Button();
            this.groupBox_payLoad = new System.Windows.Forms.GroupBox();
            this.button_testPayLoad = new System.Windows.Forms.Button();
            this.label_statusPayLoad = new System.Windows.Forms.Label();
            this.button_openPayLoad = new System.Windows.Forms.Button();
            this.groupBox_steps = new System.Windows.Forms.GroupBox();
            this.button_Step5_MultiTranToPay = new System.Windows.Forms.Button();
            this.button_step2A_siteNo = new System.Windows.Forms.Button();
            this.button_step4_TransfDataToMep = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.button_step3 = new System.Windows.Forms.Button();
            this.button_step2_missingData = new System.Windows.Forms.Button();
            this.button_step1_AddSiteNO = new System.Windows.Forms.Button();
            this.groupBoxMepStyle.SuspendLayout();
            this.groupBoxMultiTrans.SuspendLayout();
            this.groupBox_DailyTrans.SuspendLayout();
            this.groupBox_payLoad.SuspendLayout();
            this.groupBox_steps.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // groupBoxMepStyle
            // 
            this.groupBoxMepStyle.Controls.Add(this.button_clearMepStyle);
            this.groupBoxMepStyle.Controls.Add(this.buttonTestMepStyle);
            this.groupBoxMepStyle.Controls.Add(this.label_StatusMepSty);
            this.groupBoxMepStyle.Controls.Add(this.buttonOpenMepTimesheet);
            this.groupBoxMepStyle.Location = new System.Drawing.Point(12, 12);
            this.groupBoxMepStyle.Name = "groupBoxMepStyle";
            this.groupBoxMepStyle.Size = new System.Drawing.Size(855, 124);
            this.groupBoxMepStyle.TabIndex = 7;
            this.groupBoxMepStyle.TabStop = false;
            this.groupBoxMepStyle.Text = "MEP STYLE TIME SHEET";
            this.groupBoxMepStyle.Enter += new System.EventHandler(this.GroupBox1_Enter);
            // 
            // button_clearMepStyle
            // 
            this.button_clearMepStyle.Location = new System.Drawing.Point(733, 55);
            this.button_clearMepStyle.Name = "button_clearMepStyle";
            this.button_clearMepStyle.Size = new System.Drawing.Size(101, 41);
            this.button_clearMepStyle.TabIndex = 10;
            this.button_clearMepStyle.Text = "Clear";
            this.button_clearMepStyle.UseVisualStyleBackColor = true;
            this.button_clearMepStyle.Click += new System.EventHandler(this.Button_clearMepStyle_Click);
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
            this.groupBoxMultiTrans.Controls.Add(this.button_clearMultiTrans);
            this.groupBoxMultiTrans.Controls.Add(this.buttonTestMultiTrans);
            this.groupBoxMultiTrans.Controls.Add(this.label_StatusMultiTrans);
            this.groupBoxMultiTrans.Controls.Add(this.buttonOpenMultiTrans);
            this.groupBoxMultiTrans.Location = new System.Drawing.Point(12, 142);
            this.groupBoxMultiTrans.Name = "groupBoxMultiTrans";
            this.groupBoxMultiTrans.Size = new System.Drawing.Size(855, 124);
            this.groupBoxMultiTrans.TabIndex = 8;
            this.groupBoxMultiTrans.TabStop = false;
            this.groupBoxMultiTrans.Text = "MULTIPLE TRANSACTION";
            // 
            // button_clearMultiTrans
            // 
            this.button_clearMultiTrans.Location = new System.Drawing.Point(733, 55);
            this.button_clearMultiTrans.Name = "button_clearMultiTrans";
            this.button_clearMultiTrans.Size = new System.Drawing.Size(101, 41);
            this.button_clearMultiTrans.TabIndex = 10;
            this.button_clearMultiTrans.Text = "Clear";
            this.button_clearMultiTrans.UseVisualStyleBackColor = true;
            this.button_clearMultiTrans.Click += new System.EventHandler(this.Button_clearMultiTrans_Click);
            // 
            // buttonTestMultiTrans
            // 
            this.buttonTestMultiTrans.Enabled = false;
            this.buttonTestMultiTrans.Location = new System.Drawing.Point(615, 55);
            this.buttonTestMultiTrans.Name = "buttonTestMultiTrans";
            this.buttonTestMultiTrans.Size = new System.Drawing.Size(101, 41);
            this.buttonTestMultiTrans.TabIndex = 9;
            this.buttonTestMultiTrans.Text = "Test Multiple Transactions";
            this.buttonTestMultiTrans.UseVisualStyleBackColor = true;
            this.buttonTestMultiTrans.Click += new System.EventHandler(this.ButtonTestMultiTrans_Click);
            // 
            // label_StatusMultiTrans
            // 
            this.label_StatusMultiTrans.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label_StatusMultiTrans.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_StatusMultiTrans.Location = new System.Drawing.Point(246, 65);
            this.label_StatusMultiTrans.Name = "label_StatusMultiTrans";
            this.label_StatusMultiTrans.Size = new System.Drawing.Size(325, 20);
            this.label_StatusMultiTrans.TabIndex = 8;
            this.label_StatusMultiTrans.Text = "Open Multiple Transactions";
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
            this.groupBox_DailyTrans.Controls.Add(this.button_clearDailyTrans);
            this.groupBox_DailyTrans.Controls.Add(this.button_TestDailyTrans);
            this.groupBox_DailyTrans.Controls.Add(this.label_StatusDailyTrans);
            this.groupBox_DailyTrans.Controls.Add(this.button_OpenDailyTrans);
            this.groupBox_DailyTrans.Location = new System.Drawing.Point(12, 272);
            this.groupBox_DailyTrans.Name = "groupBox_DailyTrans";
            this.groupBox_DailyTrans.Size = new System.Drawing.Size(855, 124);
            this.groupBox_DailyTrans.TabIndex = 9;
            this.groupBox_DailyTrans.TabStop = false;
            this.groupBox_DailyTrans.Text = "DAILY TRANSACTIONS";
            // 
            // button_clearDailyTrans
            // 
            this.button_clearDailyTrans.Location = new System.Drawing.Point(733, 55);
            this.button_clearDailyTrans.Name = "button_clearDailyTrans";
            this.button_clearDailyTrans.Size = new System.Drawing.Size(101, 41);
            this.button_clearDailyTrans.TabIndex = 10;
            this.button_clearDailyTrans.Text = "Clear";
            this.button_clearDailyTrans.UseVisualStyleBackColor = true;
            this.button_clearDailyTrans.Click += new System.EventHandler(this.Button_clearDailyTrans_Click);
            // 
            // button_TestDailyTrans
            // 
            this.button_TestDailyTrans.Enabled = false;
            this.button_TestDailyTrans.Location = new System.Drawing.Point(615, 55);
            this.button_TestDailyTrans.Name = "button_TestDailyTrans";
            this.button_TestDailyTrans.Size = new System.Drawing.Size(101, 41);
            this.button_TestDailyTrans.TabIndex = 9;
            this.button_TestDailyTrans.Text = "Test Daily Transactions";
            this.button_TestDailyTrans.UseVisualStyleBackColor = true;
            this.button_TestDailyTrans.Click += new System.EventHandler(this.Button_TestDailyTrans_Click);
            // 
            // label_StatusDailyTrans
            // 
            this.label_StatusDailyTrans.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label_StatusDailyTrans.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_StatusDailyTrans.Location = new System.Drawing.Point(246, 65);
            this.label_StatusDailyTrans.Name = "label_StatusDailyTrans";
            this.label_StatusDailyTrans.Size = new System.Drawing.Size(325, 20);
            this.label_StatusDailyTrans.TabIndex = 8;
            this.label_StatusDailyTrans.Text = "Open Daily Transactions";
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
            // groupBox_payLoad
            // 
            this.groupBox_payLoad.Controls.Add(this.button_testPayLoad);
            this.groupBox_payLoad.Controls.Add(this.label_statusPayLoad);
            this.groupBox_payLoad.Controls.Add(this.button_openPayLoad);
            this.groupBox_payLoad.Location = new System.Drawing.Point(12, 402);
            this.groupBox_payLoad.Name = "groupBox_payLoad";
            this.groupBox_payLoad.Size = new System.Drawing.Size(855, 124);
            this.groupBox_payLoad.TabIndex = 10;
            this.groupBox_payLoad.TabStop = false;
            this.groupBox_payLoad.Text = "PAY LOAD TIMESHEET";
            // 
            // button_testPayLoad
            // 
            this.button_testPayLoad.Enabled = false;
            this.button_testPayLoad.Location = new System.Drawing.Point(615, 55);
            this.button_testPayLoad.Name = "button_testPayLoad";
            this.button_testPayLoad.Size = new System.Drawing.Size(101, 41);
            this.button_testPayLoad.TabIndex = 9;
            this.button_testPayLoad.Text = "Test PayLoad TimeSheet";
            this.button_testPayLoad.UseVisualStyleBackColor = true;
            // 
            // label_statusPayLoad
            // 
            this.label_statusPayLoad.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label_statusPayLoad.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_statusPayLoad.Location = new System.Drawing.Point(246, 65);
            this.label_statusPayLoad.Name = "label_statusPayLoad";
            this.label_statusPayLoad.Size = new System.Drawing.Size(325, 20);
            this.label_statusPayLoad.TabIndex = 8;
            this.label_statusPayLoad.Text = "Open Pay Load Time Sheet";
            // 
            // button_openPayLoad
            // 
            this.button_openPayLoad.Location = new System.Drawing.Point(27, 55);
            this.button_openPayLoad.Name = "button_openPayLoad";
            this.button_openPayLoad.Size = new System.Drawing.Size(143, 41);
            this.button_openPayLoad.TabIndex = 7;
            this.button_openPayLoad.Text = "Open and Process";
            this.button_openPayLoad.UseVisualStyleBackColor = true;
            this.button_openPayLoad.Click += new System.EventHandler(this.Button_openPayLoad_Click);
            // 
            // groupBox_steps
            // 
            this.groupBox_steps.Controls.Add(this.button_Step5_MultiTranToPay);
            this.groupBox_steps.Controls.Add(this.button_step2A_siteNo);
            this.groupBox_steps.Controls.Add(this.button_step4_TransfDataToMep);
            this.groupBox_steps.Controls.Add(this.button2);
            this.groupBox_steps.Controls.Add(this.button1);
            this.groupBox_steps.Controls.Add(this.label1);
            this.groupBox_steps.Controls.Add(this.button_step3);
            this.groupBox_steps.Controls.Add(this.button_step2_missingData);
            this.groupBox_steps.Controls.Add(this.button_step1_AddSiteNO);
            this.groupBox_steps.Location = new System.Drawing.Point(12, 532);
            this.groupBox_steps.Name = "groupBox_steps";
            this.groupBox_steps.Size = new System.Drawing.Size(855, 174);
            this.groupBox_steps.TabIndex = 11;
            this.groupBox_steps.TabStop = false;
            this.groupBox_steps.Text = "STEPS";
            // 
            // button_Step5_MultiTranToPay
            // 
            this.button_Step5_MultiTranToPay.Location = new System.Drawing.Point(662, 59);
            this.button_Step5_MultiTranToPay.Name = "button_Step5_MultiTranToPay";
            this.button_Step5_MultiTranToPay.Size = new System.Drawing.Size(123, 69);
            this.button_Step5_MultiTranToPay.TabIndex = 8;
            this.button_Step5_MultiTranToPay.Text = "STEP 5: Transfer Multi Trans Data to PayLoad";
            this.button_Step5_MultiTranToPay.UseVisualStyleBackColor = true;
            this.button_Step5_MultiTranToPay.Click += new System.EventHandler(this.Button_Step5_MultiTranToPay_Click);
            // 
            // button_step2A_siteNo
            // 
            this.button_step2A_siteNo.Location = new System.Drawing.Point(285, 59);
            this.button_step2A_siteNo.Name = "button_step2A_siteNo";
            this.button_step2A_siteNo.Size = new System.Drawing.Size(113, 69);
            this.button_step2A_siteNo.TabIndex = 7;
            this.button_step2A_siteNo.Text = "STEP 2A: Replace \'S\' in SiteNo with \'M\'";
            this.button_step2A_siteNo.UseVisualStyleBackColor = true;
            this.button_step2A_siteNo.Click += new System.EventHandler(this.Button_step2A_siteNo_Click);
            // 
            // button_step4_TransfDataToMep
            // 
            this.button_step4_TransfDataToMep.Location = new System.Drawing.Point(533, 59);
            this.button_step4_TransfDataToMep.Name = "button_step4_TransfDataToMep";
            this.button_step4_TransfDataToMep.Size = new System.Drawing.Size(123, 69);
            this.button_step4_TransfDataToMep.TabIndex = 6;
            this.button_step4_TransfDataToMep.Text = "STEP 4: Transfer MultiTrans Data to Mep Style";
            this.button_step4_TransfDataToMep.UseVisualStyleBackColor = true;
            this.button_step4_TransfDataToMep.Click += new System.EventHandler(this.Button_step4_transfDataToMep_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(763, 148);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(86, 26);
            this.button2.TabIndex = 5;
            this.button2.Text = "hide";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(693, 148);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(64, 26);
            this.button1.TabIndex = 4;
            this.button1.Text = "break";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(24, 148);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(409, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "Note: Labours Normal Working Time is Assumed as 8:00 hours;";
            // 
            // button_step3
            // 
            this.button_step3.Location = new System.Drawing.Point(404, 59);
            this.button_step3.Name = "button_step3";
            this.button_step3.Size = new System.Drawing.Size(123, 69);
            this.button_step3.TabIndex = 2;
            this.button_step3.Text = "STEP 3: Print Each Employee in Multiple Transcations";
            this.button_step3.UseVisualStyleBackColor = true;
            this.button_step3.Click += new System.EventHandler(this.Button_step3_Click);
            // 
            // button_step2_missingData
            // 
            this.button_step2_missingData.Enabled = false;
            this.button_step2_missingData.Location = new System.Drawing.Point(156, 59);
            this.button_step2_missingData.Name = "button_step2_missingData";
            this.button_step2_missingData.Size = new System.Drawing.Size(123, 69);
            this.button_step2_missingData.TabIndex = 1;
            this.button_step2_missingData.Text = "STEP 2: Add missing data from Mep Style TimeSheet to Multiple Transaction";
            this.button_step2_missingData.UseVisualStyleBackColor = true;
            this.button_step2_missingData.Click += new System.EventHandler(this.Button_step2_missingData_Click);
            // 
            // button_step1_AddSiteNO
            // 
            this.button_step1_AddSiteNO.Enabled = false;
            this.button_step1_AddSiteNO.Location = new System.Drawing.Point(27, 59);
            this.button_step1_AddSiteNO.Name = "button_step1_AddSiteNO";
            this.button_step1_AddSiteNO.Size = new System.Drawing.Size(123, 69);
            this.button_step1_AddSiteNO.TabIndex = 0;
            this.button_step1_AddSiteNO.Text = "STEP1: Add Site No. in Multiple Transaction from Daily Transaction";
            this.button_step1_AddSiteNO.UseVisualStyleBackColor = true;
            this.button_step1_AddSiteNO.Click += new System.EventHandler(this.Button_step1_AddSiteNO_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(894, 805);
            this.Controls.Add(this.groupBox_steps);
            this.Controls.Add(this.groupBox_payLoad);
            this.Controls.Add(this.groupBox_DailyTrans);
            this.Controls.Add(this.groupBoxMultiTrans);
            this.Controls.Add(this.groupBoxMepStyle);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Name = "FormMain";
            this.Text = "Attendance Handler";
            this.Activated += new System.EventHandler(this.FormMain_Activated);
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.groupBoxMepStyle.ResumeLayout(false);
            this.groupBoxMepStyle.PerformLayout();
            this.groupBoxMultiTrans.ResumeLayout(false);
            this.groupBox_DailyTrans.ResumeLayout(false);
            this.groupBox_payLoad.ResumeLayout(false);
            this.groupBox_steps.ResumeLayout(false);
            this.groupBox_steps.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
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
        private System.Windows.Forms.GroupBox groupBox_payLoad;
        private System.Windows.Forms.Button button_testPayLoad;
        private System.Windows.Forms.Label label_statusPayLoad;
        private System.Windows.Forms.Button button_openPayLoad;
        private System.Windows.Forms.GroupBox groupBox_steps;
        private System.Windows.Forms.Button button_step3;
        private System.Windows.Forms.Button button_step2_missingData;
        private System.Windows.Forms.Button button_step1_AddSiteNO;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_clearMultiTrans;
        private System.Windows.Forms.Button button_clearMepStyle;
        private System.Windows.Forms.Button button_clearDailyTrans;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button_step4_TransfDataToMep;
        private System.Windows.Forms.Button button_step2A_siteNo;
        private System.Windows.Forms.Button button_Step5_MultiTranToPay;
    }
}