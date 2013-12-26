using System.Windows.Forms;
using DevComponents.DotNetBar.Controls;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    internal class RowDept : DataGridViewRow
    {
        public RowDept(DataGridViewX dgv, string dept)
        {
            this.CreateCells(
                dgv,
                dept
            );

            this.Cells[1] = new CellGradeYear(dept, "一年級");
            this.Cells[2] = new CellGradeYear(dept, "二年級");
            this.Cells[3] = new CellGradeYear(dept, "三年級");
            this.Cells[4] = new CellGradeYear(dept, "四年級");
            this.Cells[5] = new CellGradeYear(dept, "其它");

            for (int i = 1; i <= 5; i++)
            {
                this.Cells[i].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        public void AddStudent(Student student)
        {
            switch (student.GradeYear)
            {
                case "1":
                    (this.Cells[1] as CellGradeYear).AddStudent(student);
                    break;
                case "2":
                    (this.Cells[2] as CellGradeYear).AddStudent(student);
                    break;
                case "3":
                    (this.Cells[3] as CellGradeYear).AddStudent(student);
                    break;
                case "4":
                    (this.Cells[4] as CellGradeYear).AddStudent(student);
                    break;
                default:
                    (this.Cells[5] as CellGradeYear).AddStudent(student);
                    break;
            }
        }
    }
}
