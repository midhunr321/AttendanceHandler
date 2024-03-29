﻿using System;
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
    public partial class Form_TestMepStyle : Form
    {
        public Form_TestMepStyle()
        {
            InitializeComponent();
        }

        private Boolean highlight_heading_of_mep_style_worksheet()
        {
            var worksheet = SiGlobalVars.Instance.mepStyleCurrentWorkSheet;
            if (worksheet == null)
            {
                MessageBox.Show("MEP Style Worksheet is not found in memory");
                return false;
            }
            EXCEL_HELPER eXCEL_HELPER = new EXCEL_HELPER(worksheet);

            var headings = SiGlobalVars.Instance.mepStyleHeadings;
            Color lastColour=Color.Red;

            foreach (HeadingWrap heading in headings)
            {
                
                Color color = ColourHandler.get_random_colour();
                
                eXCEL_HELPER
                    .change_cell_interior_color(ref heading.fullCell,
                    color);
            }

            foreach (var overTimedateheading in
                SiGlobalVars.Instance.mepStyleHeadings.overtimeDays)
            {
                Color color = ColourHandler.get_random_colour();
                eXCEL_HELPER
                    .change_cell_interior_color(ref overTimedateheading.Value.fullCell,
                    color);
            }
            worksheet.Activate();
            return true;
        }
        private void ButtonMepStyleHeadingTest_Click(object sender, EventArgs e)
        {
            if (highlight_heading_of_mep_style_worksheet() == false)
                return;
        }
    }
}
