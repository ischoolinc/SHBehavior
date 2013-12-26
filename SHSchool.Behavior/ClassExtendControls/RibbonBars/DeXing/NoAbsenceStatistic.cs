using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using FISCA.DSAUtil;
using System.Xml;
using Aspose.Cells;
using System.IO;
using DevComponents.DotNetBar;
using System.Diagnostics;
using SmartSchool.Common;
using K12.Data;
using SHSchool.Behavior.StudentExtendControls;

namespace SHSchool.Behavior.ClassExtendControls
{
    public partial class NoAbsenceStatistic : UserControl, IDeXingExport
    {
        private string[] _classidList;
        private string _schoolYear;
        private string _semester;

        public NoAbsenceStatistic(string[] classidList)
        {
            InitializeComponent();
            _classidList = classidList;
        }

        #region IDeXingExport 成員

        public Control MainControl
        {
            get { return this.groupPanel1; }
        }

        public void LoadData()
        {
            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");
            cboSemester.SelectedIndex = int.Parse(School.DefaultSemester) - 1;

            int schoolYear = int.Parse(School.DefaultSchoolYear);
            for (int i = schoolYear; i > schoolYear - 4; i--)
            {
                cboSchoolYear.Items.Add(i);
            }
            if (cboSchoolYear.Items.Count > 0)
                cboSchoolYear.SelectedIndex = 0;

            //_schoolYear = SmartSchool.Common.CurrentUser.Instance.SchoolYear.ToString();
            //_semester = SmartSchool.Common.CurrentUser.Instance.Semester.ToString();

            //checkBoxX1.Text = checkBoxX1.Text.Replace("@@", _schoolYear).Replace("!!", _semester);
            checkBoxX1.Checked = true;
        }

        public void Export()
        {
            //取得班級ID
            List<string> ClassIDList = K12.Presentation.NLDPanels.Class.SelectedSource;
            List<string> 會影響全勤的假別 = new List<string>();
            foreach (AbsenceMappingInfo each in AbsenceMapping.SelectAll())
            {
                if (!each.Noabsence)
                {
                    會影響全勤的假別.Add(each.Name);
                }
            }

            //取得班級學生
            List<StudentRecord> StudentList = Student.SelectByClassIDs(ClassIDList);

            //所有學生
            Dictionary<string, StudentRecord> StudentDic = new Dictionary<string, StudentRecord>();
            //不全勤學生
            Dictionary<string, StudentRecord> noOK_StudentList = new Dictionary<string, StudentRecord>();
            //全勤的學生
            Dictionary<string, StudentRecord> OK_StudentList = new Dictionary<string, StudentRecord>();

            foreach (StudentRecord each in StudentList)
            {
                if (each.Status == StudentRecord.StudentStatus.一般 || each.Status == StudentRecord.StudentStatus.延修)
                {
                    if (!StudentDic.ContainsKey(each.ID))
                    {
                        StudentDic.Add(each.ID, each);
                    }
                }
            }

            //取得學生所有的缺曠資料
            //如果沒有缺曠記錄的
            //就算是全勤
            List<AttendanceRecord> TestList = Attendance.SelectByStudentIDs(StudentDic.Keys);
            List<AttendanceRecord> AttendanceList = new List<AttendanceRecord>();

            if (checkBoxX1.Checked)
            {
                //相同學年度/學期
                foreach (AttendanceRecord attendnace in TestList)
                {
                    int schoolYear = int.Parse(cboSchoolYear.SelectedItem.ToString());
                    int semester = int.Parse(cboSemester.SelectedItem.ToString());
                    if (attendnace.SchoolYear == schoolYear && attendnace.Semester == semester)
                    {
                        AttendanceList.Add(attendnace);
                    }
                }
            }
            else
            {
                AttendanceList.AddRange(TestList);
            }

            foreach (AttendanceRecord each in AttendanceList)
            {
                foreach (K12.Data.AttendancePeriod period in each.PeriodDetail)
                {
                    if (會影響全勤的假別.Contains(period.AbsenceType))
                    {
                        if (!noOK_StudentList.ContainsKey(period.RefStudentID))
                        {
                            noOK_StudentList.Add(period.RefStudentID, StudentDic[period.RefStudentID]);
                        }
                    }
                }
            }

            foreach (StudentRecord each in StudentDic.Values)
            {
                if (!noOK_StudentList.ContainsKey(each.ID))
                {
                    if (!OK_StudentList.ContainsKey(each.ID))
                    {
                        OK_StudentList.Add(each.ID, each);
                    }
                }
            }
            List<StudentRecord> StudentRecordList = new List<StudentRecord>();
            foreach (StudentRecord each in OK_StudentList.Values)
            {
                StudentRecordList.Add(each);
            }

            StudentRecordList = SortClassIndex.K12Data_StudentRecord(StudentRecordList);

            Workbook book = new Workbook();
            Worksheet sheet = book.Worksheets[0];

            string schoolName = School.ChineseName;
            Cell A1 = sheet.Cells["A1"];
            A1.Style.Borders.SetColor(Color.Black);
            string A1Name = schoolName + "  ";
            if (checkBoxX1.Checked)
            {
                A1Name += "(" + cboSchoolYear.SelectedItem.ToString() + "/" + cboSemester.SelectedItem.ToString() + ") ";
            }

            A1Name += "全勤學生名單";
            sheet.Name = "全勤學生名單";
            A1.PutValue(A1Name);
            A1.Style.HorizontalAlignment = TextAlignmentType.Center;
            sheet.Cells.Merge(0, 0, 1, 5);

            FormatCell(sheet.Cells["A2"], "編號");
            FormatCell(sheet.Cells["B2"], "班級");
            FormatCell(sheet.Cells["C2"], "座號");
            FormatCell(sheet.Cells["D2"], "姓名");
            FormatCell(sheet.Cells["E2"], "學號");

            int index = 1;
            foreach (StudentRecord e in StudentRecordList)
            {
                int rowIndex = index + 2;
                FormatCell(sheet.Cells["A" + rowIndex], index.ToString());
                FormatCell(sheet.Cells["B" + rowIndex], string.IsNullOrEmpty(e.RefClassID) ? "" : e.Class.Name);
                FormatCell(sheet.Cells["C" + rowIndex], e.SeatNo.HasValue ? e.SeatNo.Value.ToString() : "");
                FormatCell(sheet.Cells["D" + rowIndex], e.Name);
                FormatCell(sheet.Cells["E" + rowIndex], e.StudentNumber);
                index++;
            }
            string path = Path.Combine(Application.StartupPath, "Reports");
            path = Path.Combine(path, book.Worksheets[0].Name + ".xls");
            int i = 1;
            while (true)
            {
                string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                if (!File.Exists(newPath))
                {
                    path = newPath;
                    break;
                }
            }
            try
            {
                book.Save(path);
            }
            catch (IOException)
            {
                try
                {
                    FileInfo file = new FileInfo(path);
                    string nameTempalte = file.FullName.Replace(file.Extension, "") + "{0}.xls";
                    int count = 1;
                    string fileName = string.Format(nameTempalte, count);
                    while (File.Exists(fileName))
                        fileName = string.Format(nameTempalte, count++);

                    book.Save(fileName);
                    path = fileName;
                }
                catch (Exception ex)
                {
                    MsgBox.Show("檔案儲存失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show("檔案儲存失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Process.Start(path);
            }
            catch (Exception ex)
            {
                MsgBox.Show("檔案開啟失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void FormatCell(Cell cell, string value)
        {
            cell.PutValue(value);
            cell.Style.Borders.SetStyle(CellBorderType.Hair);
            cell.Style.Borders.SetColor(Color.Black);
            cell.Style.Borders.DiagonalStyle = CellBorderType.None;
            cell.Style.HorizontalAlignment = TextAlignmentType.Center;
        }

        #endregion


    }
}
