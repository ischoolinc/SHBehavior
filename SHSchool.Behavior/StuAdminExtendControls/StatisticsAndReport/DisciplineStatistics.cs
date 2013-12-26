using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Xml;
using DevComponents.DotNetBar;
using FISCA.DSAUtil;
using K12.Data.Utility;
using SmartSchool.Customization.Data;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    public class DisciplineStatistics
    {
        private BackgroundWorker _disciplineLoader;
        private ManualResetEvent _wait;

        private DisciplineCollection _disciplines;
        private StudentCollection _students;

        private int _school_year;
        private int _semester;

        private DistributeForm _distributeForm;

        public DisciplineStatistics()
        {
            SelectSemesterForm form = new SelectSemesterForm();
            if (form.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            _disciplines = new DisciplineCollection();
            _students = new StudentCollection();

            _school_year = form.SchoolYear;
            _semester = form.Semester;

            _wait = new ManualResetEvent(false);

            _disciplineLoader = new BackgroundWorker();
            _disciplineLoader.DoWork += new DoWorkEventHandler(_disciplineLoader_DoWork);
            _disciplineLoader.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_disciplineLoader_RunWorkerCompleted);
            _disciplineLoader.ProgressChanged += new ProgressChangedEventHandler(_disciplineLoader_ProgressChanged);
            _disciplineLoader.WorkerReportsProgress = true;
            _disciplineLoader.RunWorkerAsync();
        }

        private List<Student> GetErrorStudentGender()
        {
            List<Student> error_list = new List<Student>();

            lock (_students)
            {
                foreach (Student stu in _students.Values)
                {
                    if (stu.Gender != "男" && stu.Gender != "女")
                        error_list.Add(stu);
                }
            }

            return error_list;
        }

        private void _disciplineLoader_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取學生獎懲資料中...", e.ProgressPercentage);
        }

        private void _disciplineLoader_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取完成");

            List<Student> error = GetErrorStudentGender();
            if (error.Count > 0)
            {
                ErrorViewer viewer = new ErrorViewer(error);
                viewer.ShowInTaskbar = true;
                viewer.Show();

                return;
            }
            _distributeForm = new DistributeForm(_students, _school_year.ToString(), _semester.ToString());
            if (_distributeForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ReportGenerator generator = new ReportGenerator(_students, _school_year.ToString(), _semester.ToString());
                string location = generator.Generate();

                if (string.IsNullOrEmpty(location))
                {
                    MessageBoxEx.Show("儲存失敗");
                    return;
                }

                System.Diagnostics.Process.Start(location);
            }
        }

        private void _disciplineLoader_DoWork(object sender, DoWorkEventArgs e)
        {
            _disciplineLoader.ReportProgress(0);

            _students.Clear();
            Dictionary<string, List<Discipline>> studentDisciplines = new Dictionary<string, List<Discipline>>();

            #region 抓所有獎懲紀錄

            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("Field");
            helper.AddElement("Field", "All");
            helper.AddElement("Condition");
            helper.AddElement("Condition", "SchoolYear", _school_year.ToString());
            helper.AddElement("Condition", "Semester", _semester.ToString());
            DSResponse dsrsp = DSAServices.CallService("SmartSchool.Student.Discipline.GetDiscipline", new DSRequest(helper));

            double dis_amount = dsrsp.GetContent().GetElements("Discipline").Length;
            double dis_count = 0;

            foreach (XmlElement disElement in dsrsp.GetContent().GetElements("Discipline"))
            {
                dis_count++;

                DSXmlHelper disHelper = new DSXmlHelper(disElement);

                string id = disHelper.GetText("RefStudentID");
                string kind = "";
                switch (disHelper.GetText("MeritFlag"))
                {
                    case "1":
                        kind = "獎勵";
                        break;
                    case "0":
                        kind = "懲罰";
                        break;
                    case "2":
                        kind = "留校察看";
                        break;
                    default:
                        break;
                }

                Discipline dis = new Discipline(id, kind);

                if (disHelper.GetText("MeritFlag") == "1")
                {
                    dis.AwardA = int.Parse(disHelper.GetText("Detail/Discipline/Merit/@A"));
                    dis.AwardB = int.Parse(disHelper.GetText("Detail/Discipline/Merit/@B"));
                    dis.AwardC = int.Parse(disHelper.GetText("Detail/Discipline/Merit/@C"));
                }

                if (disHelper.GetText("MeritFlag") == "0")
                {
                    dis.FaultA = int.Parse(disHelper.GetText("Detail/Discipline/Demerit/@A"));
                    dis.FaultB = int.Parse(disHelper.GetText("Detail/Discipline/Demerit/@B"));
                    dis.FaultC = int.Parse(disHelper.GetText("Detail/Discipline/Demerit/@C"));

                    dis.Cleared = (disHelper.GetText("Detail/Discipline/Demerit/@Cleared") == "是") ? true : false;
                }
                if (!dis.Cleared)
                {
                    if (!studentDisciplines.ContainsKey(dis.StudentID))
                        studentDisciplines.Add(dis.StudentID, new List<Discipline>());
                    studentDisciplines[dis.StudentID].Add(dis);
                }

                _disciplineLoader.ReportProgress((int)((dis_count * 30) / dis_amount));
            }

            #endregion

            AccessHelper accessHelper = new AccessHelper();

            //抓有獎懲的學生
            List<StudentRecord> studentRecs = accessHelper.StudentHelper.GetStudents(studentDisciplines.Keys);

            double stu_amount = studentRecs.Count;
            double stu_count = 0;

            #region 填入學生學期歷程(作廢)
            //填入學生學期歷程
            //List<StudentRecord> partList = new List<StudentRecord>();
            //int size = 50;
            //foreach (StudentRecord stu in studentRecs)
            //{
            //    stu_count++;

            //    partList.Add(stu);
            //    if (partList.Count == size)
            //    {
            //        accessHelper.StudentHelper.FillSemesterHistory(partList);
            //        partList.Clear();
            //    }

            //    _disciplineLoader.ReportProgress(30 + (int)((stu_count * 50) / stu_amount));
            //}
            //if (partList.Count > 0)
            //    accessHelper.StudentHelper.FillSemesterHistory(partList);
            #endregion

            #region 整理成好用的學生資料

            stu_count = 0;

            foreach (StudentRecord studentRec in studentRecs)
            {
                stu_count++;

                Student stu = new Student(studentRec.StudentID);
                stu.Name = studentRec.StudentName;
                stu.StudentNumber = studentRec.StudentNumber;
                stu.Gender = studentRec.Gender;
                stu.Department = studentRec.Department;
                stu.ClassName = studentRec.RefClass != null ? studentRec.RefClass.ClassName : "";
                stu.GradeYear = studentRec.RefClass != null ? studentRec.RefClass.GradeYear : "";

                // 作廢
                //foreach (SemesterHistory his in studentRec.SemesterHistoryList)
                //{
                //    if (his.SchoolYear == _school_year && his.Semester == _semester)
                //    {
                //        stu.GradeYear = "" + his.GradeYear;
                //        break;
                //    }
                //}

                foreach (Discipline dis in studentDisciplines[studentRec.StudentID])
                {
                    if (dis.Kind == "留校察看")
                        stu.UltimateAdmonition = true;
                    stu.AwardA += dis.AwardA;
                    stu.AwardB += dis.AwardB;
                    stu.AwardC += dis.AwardC;
                    stu.FaultA += dis.FaultA;
                    stu.FaultB += dis.FaultB;
                    stu.FaultC += dis.FaultC;
                }

                _students.Add(stu.ID, stu);

                _disciplineLoader.ReportProgress(80 + (int)((stu_count * 20) / stu_amount));
            }

            #endregion
        }
    }
}
