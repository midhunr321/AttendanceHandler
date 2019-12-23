namespace AttendanceHander
{
    partial class Form_holidaysSelector
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
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.button_addHoliday = new System.Windows.Forms.Button();
            this.button_removeHoliday = new System.Windows.Forms.Button();
            this.button_OK = new System.Windows.Forms.Button();
            this.dataGridViewHolidays = new System.Windows.Forms.DataGridView();
            this.Holidays = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewHolidays)).BeginInit();
            this.SuspendLayout();
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(37, 21);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(379, 20);
            this.dateTimePicker1.TabIndex = 0;
            // 
            // button_addHoliday
            // 
            this.button_addHoliday.Location = new System.Drawing.Point(71, 74);
            this.button_addHoliday.Name = "button_addHoliday";
            this.button_addHoliday.Size = new System.Drawing.Size(142, 37);
            this.button_addHoliday.TabIndex = 2;
            this.button_addHoliday.Text = "Add Holiday Date";
            this.button_addHoliday.UseVisualStyleBackColor = true;
            this.button_addHoliday.Click += new System.EventHandler(this.Button_addHoliday_Click);
            // 
            // button_removeHoliday
            // 
            this.button_removeHoliday.Location = new System.Drawing.Point(237, 74);
            this.button_removeHoliday.Name = "button_removeHoliday";
            this.button_removeHoliday.Size = new System.Drawing.Size(146, 37);
            this.button_removeHoliday.TabIndex = 3;
            this.button_removeHoliday.Text = "Remove Selected from List";
            this.button_removeHoliday.UseVisualStyleBackColor = true;
            this.button_removeHoliday.Click += new System.EventHandler(this.Button_removeHoliday_Click);
            // 
            // button_OK
            // 
            this.button_OK.Location = new System.Drawing.Point(149, 343);
            this.button_OK.Name = "button_OK";
            this.button_OK.Size = new System.Drawing.Size(146, 36);
            this.button_OK.TabIndex = 4;
            this.button_OK.Text = "OK";
            this.button_OK.UseVisualStyleBackColor = true;
            this.button_OK.Click += new System.EventHandler(this.Button_OK_Click);
            // 
            // dataGridViewHolidays
            // 
            this.dataGridViewHolidays.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewHolidays.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Holidays});
            this.dataGridViewHolidays.Location = new System.Drawing.Point(37, 141);
            this.dataGridViewHolidays.MultiSelect = false;
            this.dataGridViewHolidays.Name = "dataGridViewHolidays";
            this.dataGridViewHolidays.Size = new System.Drawing.Size(379, 133);
            this.dataGridViewHolidays.TabIndex = 5;
            this.dataGridViewHolidays.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridViewHolidays_CellContentClick);
            // 
            // Holidays
            // 
            this.Holidays.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Holidays.HeaderText = "Holidays";
            this.Holidays.Name = "Holidays";
            // 
            // Form_holidaysSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(455, 391);
            this.Controls.Add(this.dataGridViewHolidays);
            this.Controls.Add(this.button_OK);
            this.Controls.Add(this.button_removeHoliday);
            this.Controls.Add(this.button_addHoliday);
            this.Controls.Add(this.dateTimePicker1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form_holidaysSelector";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Holidays Selector";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Form_holidaysSelector_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewHolidays)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Button button_addHoliday;
        private System.Windows.Forms.Button button_removeHoliday;
        private System.Windows.Forms.Button button_OK;
        private System.Windows.Forms.DataGridView dataGridViewHolidays;
        private System.Windows.Forms.DataGridViewTextBoxColumn Holidays;
    }
}