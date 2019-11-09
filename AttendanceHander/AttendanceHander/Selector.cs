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
    public partial class Selector : Form
    {
        String label;
        List<String> dataLists;
        Form rootForm;
        public String selected_employeePosition { get; set; }

        public Selector(String label,List<String> dataLists,
            Form rootForm)
        {
            InitializeComponent();
            this.label = label;
            this.dataLists = dataLists;
            this.rootForm = rootForm;
        }

        private void Button_Ok_Click(object sender, EventArgs e)
        {
            selected_employeePosition = comboBox1.SelectedItem.ToString();
            this.Hide();

            this.DialogResult = DialogResult.OK;
            rootForm.Activate();
        }

        private void Selector_Load(object sender, EventArgs e)
        {
            update_combobox_with_worksheets();
        }

      

        private void update_combobox_with_worksheets()
        {
           foreach(var list in dataLists)
            {
                comboBox1.Items.Add(list);
            }

            comboBox1.SelectedIndex = 0;
        }
    }
}
