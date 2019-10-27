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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.buttonTestMepStyle = new System.Windows.Forms.Button();
            this.labelMepStyle = new System.Windows.Forms.Label();
            this.buttonOpenMepTimesheet = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
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
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.buttonTestMepStyle);
            this.groupBox1.Controls.Add(this.labelMepStyle);
            this.groupBox1.Controls.Add(this.buttonOpenMepTimesheet);
            this.groupBox1.Location = new System.Drawing.Point(12, 121);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(722, 124);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "MEP STYLE TIME SHEET";
            this.groupBox1.Enter += new System.EventHandler(this.GroupBox1_Enter);
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
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.button4);
            this.groupBox2.Location = new System.Drawing.Point(12, 251);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(722, 124);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "MULTIPLE TRANSACTION";
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.Location = new System.Drawing.Point(615, 55);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(101, 41);
            this.button2.TabIndex = 9;
            this.button2.Text = "Test MEP Style Timesheet";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(246, 65);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(325, 20);
            this.label1.TabIndex = 8;
            this.label1.Text = "Open Plumbers Time Sheet (Default MEP Style)";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(27, 55);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(143, 41);
            this.button4.TabIndex = 7;
            this.button4.Text = "Open and Process";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(800, 514);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Name = "FormMain";
            this.Text = "Attendance Handler";
            this.Activated += new System.EventHandler(this.FormMain_Activated);
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button buttonTestMepStyle;
        private System.Windows.Forms.Label labelMepStyle;
        private System.Windows.Forms.Button buttonOpenMepTimesheet;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button4;
    }
}