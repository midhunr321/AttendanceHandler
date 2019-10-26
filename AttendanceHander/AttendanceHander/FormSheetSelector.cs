using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace AttendanceHander
{
    public partial class FormSheetSelector : Form
    {
        Excel.Workbook workbook;
        Form previousForm;
        public FormSheetSelector(Excel.Workbook workbook, 
            Form previousForm)
        {
            this.workbook = workbook;
            this.previousForm = previousForm;
            InitializeComponent();
        }

        private void ButtonOk_Click(object sender, EventArgs e)
        {
            Excel.Worksheet currentMonthSheet
                = ((KeyValuePair<string, Excel.Worksheet>)comboBoxSheets
                .SelectedItem).Value;

            SiGlobalVars.Instance.mepStyleCurrentWorkSheet = currentMonthSheet;
            this.Hide();
            this.DialogResult = DialogResult.OK;
            previousForm.Activate();

        }

        private void update_combobox_with_worksheets(Excel.Sheets sheets)
        {
            Dictionary<String, Excel.Worksheet> keyValuePairs
                   = new Dictionary<string, Excel.Worksheet>();
            foreach (Excel.Worksheet sheet in sheets)
            {
                keyValuePairs.Add(sheet.Name, sheet);
            }
            comboBoxSheets.DataSource = new BindingSource(keyValuePairs, null);
            comboBoxSheets.DisplayMember = "Key";
            comboBoxSheets.ValueMember = "Value";
            int last = comboBoxSheets.Items.Count;
            comboBoxSheets.SelectedIndex = last-1;
        }
        private void FormSheetSelector_Load(object sender, EventArgs e)
        {
            if (workbook == null)
            {
                MessageBox.Show("Workbook is null");
                this.Hide();
                return;
            }

            Excel.Sheets sheets = workbook.Sheets;
            update_combobox_with_worksheets(sheets);

        }

        private void FormSheetSelector_FormClosed(object sender, FormClosedEventArgs e)
        {
            previousForm.Activate();
        }
    }
}
