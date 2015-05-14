using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;

namespace SHSchool.Behavior
{
    public partial class SelectSchoolYearSemesterForm : SmartSchool.Common.BaseForm
    {
        public string SchoolYear
        {
            get
            {
                if (checkBoxX1.Checked)
                    return "";
                else
                    return numSchoolYear.Value.ToString();
            }
        }

        public string Semester
        {
            get
            {
                if (checkBoxX1.Checked)
                    return "";
                else
                    return numSemester.Value.ToString();
            }
        }

        public SelectSchoolYearSemesterForm()
        {
            InitializeComponent();

            numSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            numSemester.Value = int.Parse(K12.Data.School.DefaultSemester);

        }

        protected virtual void buttonX1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxX1.Checked)
            {
                numSchoolYear.Enabled = false;
                numSemester.Enabled = false;
            }
            else
            {
                numSchoolYear.Enabled = true;
                numSemester.Enabled = true;
            }
        }
    }
}