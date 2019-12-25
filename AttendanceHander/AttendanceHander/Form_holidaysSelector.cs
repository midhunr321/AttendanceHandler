using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceHander
{
    public partial class Form_holidaysSelector : Form
    {

        List<DateTime> selectedHolidays;
        private Form previousForm;

        public Form_holidaysSelector(Form previousForm)
        {
            this.previousForm = previousForm;
            InitializeComponent();
        }

        public List<DateTime> SelectedHolidays { get => selectedHolidays;  }

        private void bind_dataGridViewHolidays()
        {
            var bindSource = new BindingSource();
            bindSource.DataSource = selectedHolidays;
            dataGridViewHolidays.DataSource = bindSource;

        }
        private void Button_addHoliday_Click(object sender, EventArgs e)
        {

            DateTime selectedDate = dateTimePicker1.Value.Date;
            if (selectedHolidays == null)
                selectedHolidays = new List<DateTime>();

            selectedHolidays.Add(selectedDate);
            bind_dataGridViewHolidays();
            dataGridViewHolidays.Invalidate();
            
        }

        private void Form_holidaysSelector_Load(object sender, EventArgs e)
        {
            if (selectedHolidays == null)
                selectedHolidays = new List<DateTime>();
            bind_dataGridViewHolidays();
            dataGridViewHolidays.Invalidate();
        }

        private void Button_removeHoliday_Click(object sender, EventArgs e)
        {
            if (dataGridViewHolidays.CurrentRow != null)
            {
                selectedHolidays.RemoveAt(dataGridViewHolidays.CurrentRow.Index);
                bind_dataGridViewHolidays();
                dataGridViewHolidays.Invalidate();
            }
        }

        private void DataGridViewHolidays_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Button_OK_Click(object sender, EventArgs e)
        {

            this.Hide();
            this.DialogResult = DialogResult.OK;
            if(previousForm!=null)
            previousForm.Activate();
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
