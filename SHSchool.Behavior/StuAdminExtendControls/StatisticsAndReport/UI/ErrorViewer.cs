using System;
using System.Collections.Generic;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    internal partial class ErrorViewer : BaseForm
    {
        private List<Student> _list;

        public ErrorViewer(List<Student> list)
        {
            InitializeComponent();
            _list = list;
        }

        private void ErrorViewer_Load(object sender, EventArgs e)
        {
            foreach (Student stu in _list)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1, stu.ClassName, stu.StudentNumber, stu.Name, "沒有性別");
                dataGridViewX1.Rows.Add(row);
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}