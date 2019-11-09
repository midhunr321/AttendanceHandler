namespace AttendanceHander
{
    partial class Form_ExportStatus
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
            this.button_cancelExport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button_cancelExport
            // 
            this.button_cancelExport.Location = new System.Drawing.Point(213, 136);
            this.button_cancelExport.Name = "button_cancelExport";
            this.button_cancelExport.Size = new System.Drawing.Size(146, 28);
            this.button_cancelExport.TabIndex = 0;
            this.button_cancelExport.Text = "Cancel Exporting";
            this.button_cancelExport.UseVisualStyleBackColor = true;
            this.button_cancelExport.Click += new System.EventHandler(this.Button_cancelExport_Click);
            // 
            // Form_ExportStatus
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(549, 198);
            this.Controls.Add(this.button_cancelExport);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.Name = "Form_ExportStatus";
            this.Text = "Exporting";
            this.Load += new System.EventHandler(this.Form_ExportStatus_Load);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.Button button_cancelExport;
    }
}