using System;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    public partial class SelectSemesterForm : BaseForm
    {
        private int _school_year = 50;
        public int SchoolYear
        {
            get { return _school_year; }
        }

        private int _semester = 1;
        public int Semester
        {
            get { return _semester; }
        }

        public SelectSemesterForm()
        {
            InitializeComponent();
        }

        private void SelectSemesterForm_Load(object sender, EventArgs e)
        {
            numericUpDown1.Value = int.Parse(School.DefaultSchoolYear);
            numericUpDown2.Value = int.Parse(School.DefaultSemester);
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            _school_year = (int)numericUpDown1.Value;
            _semester = (int)numericUpDown2.Value;

            this.DialogResult = DialogResult.OK;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}