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
            this.buttonOpenMepTimesheet = new System.Windows.Forms.Button();
            this.labelMepStyle = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.buttonTestMepStyle = new System.Windows.Forms.Button();
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
            // buttonOpenMepTimesheet
            // 
            this.buttonOpenMepTimesheet.Location = new System.Drawing.Point(57, 146);
            this.buttonOpenMepTimesheet.Name = "buttonOpenMepTimesheet";
            this.buttonOpenMepTimesheet.Size = new System.Drawing.Size(143, 41);
            this.buttonOpenMepTimesheet.TabIndex = 1;
            this.buttonOpenMepTimesheet.Text = "Open and Process";
            this.buttonOpenMepTimesheet.UseVisualStyleBackColor = true;
            this.buttonOpenMepTimesheet.Click += new System.EventHandler(this.ButtonOpenMepStyle_Click);
            // 
            // labelMepStyle
            // 
            this.labelMepStyle.AutoSize = true;
            this.labelMepStyle.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelMepStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelMepStyle.Location = new System.Drawing.Point(217, 146);
            this.labelMepStyle.Name = "labelMepStyle";
            this.labelMepStyle.Size = new System.Drawing.Size(325, 20);
            this.labelMepStyle.TabIndex = 2;
            this.labelMepStyle.Text = "Open Plumbers Time Sheet (Default MEP Style)";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(217, 276);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 18);
            this.label2.TabIndex = 4;
            this.label2.Text = "label2";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(57, 266);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(143, 41);
            this.button3.TabIndex = 3;
            this.button3.Text = "Open";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // buttonTestMepStyle
            // 
            this.buttonTestMepStyle.Enabled = false;
            this.buttonTestMepStyle.Location = new System.Drawing.Point(560, 146);
            this.buttonTestMepStyle.Name = "buttonTestMepStyle";
            this.buttonTestMepStyle.Size = new System.Drawing.Size(101, 41);
            this.buttonTestMepStyle.TabIndex = 5;
            this.buttonTestMepStyle.Text = "Test MEP Style Timesheet";
            this.buttonTestMepStyle.UseVisualStyleBackColor = true;
            this.buttonTestMepStyle.Click += new System.EventHandler(this.ButtonTestMepStyle_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(800, 514);
            this.Controls.Add(this.buttonTestMepStyle);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.labelMepStyle);
            this.Controls.Add(this.buttonOpenMepTimesheet);
            this.Controls.Add(this.button1);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Name = "FormMain";
            this.Text = "Attendance Handler";
            this.Activated += new System.EventHandler(this.FormMain_Activated);
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button buttonOpenMepTimesheet;
        private System.Windows.Forms.Label labelMepStyle;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button buttonTestMepStyle;
    }
}