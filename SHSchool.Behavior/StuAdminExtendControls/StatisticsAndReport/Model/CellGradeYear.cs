using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    internal class CellGradeYear : DataGridViewTextBoxCell
    {
        private DetailForm _detailForm;

        public CellGradeYear(string dept, string grade_year)
        {
            _dept = dept;
            _grade_year = grade_year;
            _students = new StudentCollection();
        }

        public DetailForm Detail
        {
            get
            {
                if (_detailForm == null)
                    _detailForm = new DetailForm(this);
                _detailForm.Students = _students;
                return _detailForm;
            }
        }

        public void AddStudent(Student student)
        {
            if (!_students.ContainsKey(student.ID))
                _students.Add(student.ID, student);

            this.Value = Count + "/" + Amount;
        }

        public void SetCategory(List<Student> students, string cate)
        {
            if (students != null)
            {
                foreach (Student stu in students)
                {
                    if (cate == "")
                    {
                        stu.Location = "";
                    }
                    else if (!cate.Contains("/"))
                    {
                        if (!string.IsNullOrEmpty(stu.GradeYear))
                        {
                            string gy = "";
                            if (stu.GradeYear == "1")
                                gy = "一年級";
                            else if (stu.GradeYear == "2")
                                gy = "二年級";
                            else if (stu.GradeYear == "3")
                                gy = "三年級";
                            else if (stu.GradeYear == "4")
                                gy = "四年級";

                            if (!string.IsNullOrEmpty(gy))
                                stu.Location = cate + "/" + gy;
                        }
                        else
                            continue;
                    }
                    else
                        stu.Location = cate;
                }
            }

            if (this.Count == this.Amount)
                this.Style.BackColor = Color.GreenYellow;
            else if(this.Count > 0)
                this.Style.BackColor = Color.Yellow;
            else
                this.Style.BackColor = Color.White;

            this.Value = Count + "/" + Amount;
        }

        private StudentCollection _students;
        public StudentCollection Students
        {
            get { return _students; }
        }

        private string _dept;
        public string Department
        {
            get { return _dept; }
        }

        private string _grade_year;
        public string GradeYear
        {
            get { return _grade_year; }
        }

        public int Amount
        {
            get
            {
                if (_students != null)
                    return _students.Count;
                return 0;
            }
        }

        public int Count
        {
            get
            {
                if (_students != null)
                {
                    int count = 0;
                    foreach (Student stu in _students.Values)
                    {
                        if (!string.IsNullOrEmpty(stu.Location))
                            count++;
                    }
                    return count;
                }
                return 0;
            }
        }

        //public new object Value
        //{
        //    get
        //    {
        //        return Count + "/" + Amount;
        //    }
        //    set { Value = value; }
        //}
    }
}
