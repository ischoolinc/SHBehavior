using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FISCA.DSAUtil;
using System.Xml;
using DevComponents.DotNetBar;
using FISCA.LogAgent;
using SHSchool.Behavior.StuAdminExtendControls;
using SHSchool.Data;
using Framework.Feature;
using K12.Data;
using SHSchool.Behavior.Feature;
using FISCA.Presentation.Controls;

namespace SHSchool.Behavior.StudentExtendControls
{
    public partial class TestSingleEditor : FISCA.Presentation.Controls.BaseForm
    {
        private AbsenceInfo _checkedAbsence;
        private Dictionary<string, AbsenceInfo> _absenceList;
        private List<string> _studentList;
        private ISemester _semesterProvider;
        private int _startIndex;
        private ErrorProvider _errorProvider;
        private DateTime _currentStartDate;
        private DateTime _currentEndDate;

        List<DataGridViewRow> _hiddenRows; //���ê�Rows

        List<string> WeekDay = new List<string>();
        List<DayOfWeek> nowWeekDay = new List<DayOfWeek>();

        Dictionary<string, int> ColumnIndex = new Dictionary<string, int>();

        //System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();

        //log �ݭn�Ψ쪺
        //<�ǥ�,<���<�`��,���m�O>>
        //<�ǥ�ID,�C�@��Log�O��>
        private Dictionary<string, LogStudent> LOG = new Dictionary<string, LogStudent>();

        public TestSingleEditor(List<string> studentList)
        {
            InitializeComponent(); //�]�p�u�㲣�ͪ�

            _errorProvider = new ErrorProvider();
            _studentList = studentList;
            _absenceList = new Dictionary<string, AbsenceInfo>();
            _semesterProvider = SemesterProvider.GetInstance(); 

            _hiddenRows = new List<DataGridViewRow>();
        }

        private void SingleEditor_Load(object sender, EventArgs e)
        {
            #region Load
            this.Text = "�h�H�����n��";
            InitializeRadioButton(); //���m���O�إ�
            InitializeDateRange(); //���o����w�q
            InitializeDataGridViewColumn(); //DataGridView��Column�إ�
            //SearchDateRange();
            //GetAbsense();
            LoadAbsense();

            #endregion
        }

        private void InitializeDateRange()
        {
            #region ����w�q
            K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["Attendance_BatchEditor_ByMany"];

            string date = DateConfig["SingleEditor"];

            if (date == "")
            {
                DSXmlHelper helper = new DSXmlHelper("xml");
                helper.AddElement("StartDate");
                helper.AddText("StartDate", DateTime.Today.AddDays(-6).ToShortDateString());
                helper.AddElement("EndDate");
                helper.AddText("EndDate", DateTime.Today.ToShortDateString());
                helper.AddElement("Locked");
                helper.AddText("Locked", "false");

                date = helper.BaseElement.OuterXml;
                DateConfig["SingleEditor"] = date;
                DateConfig.Save(); //�x�s���w�]��
            }

            XmlElement loadXml = DSXmlHelper.LoadXml(date);
            checkBoxX1.Checked = bool.Parse(loadXml.SelectSingleNode("Locked").InnerText);

            if (checkBoxX1.Checked) //�p�G�O��w,�N����w���
            {
                dateTimeInput1.Text = loadXml.SelectSingleNode("StartDate").InnerText;
                dateTimeInput2.Text = loadXml.SelectSingleNode("EndDate").InnerText;
            }
            else //�p�G�S����w,�N�����
            {
                dateTimeInput1.Text = DateTime.Today.AddDays(-6).ToShortDateString();
                dateTimeInput2.Text = DateTime.Today.ToShortDateString();
            }
            _currentStartDate = dateTimeInput1.Value;
            _currentEndDate = dateTimeInput2.Value;
            #endregion
        }

        private void SaveDateSetting()
        {
            #region �x�s������
            K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["Attendance_BatchEditor_ByMany"];

            DSXmlHelper helper = new DSXmlHelper("xml");
            helper.AddElement("StartDate");
            helper.AddText("StartDate", dateTimeInput1.Value.ToShortDateString());
            helper.AddElement("EndDate");
            helper.AddText("EndDate", dateTimeInput2.Value.ToShortDateString());
            helper.AddElement("Locked");
            helper.AddText("Locked", checkBoxX1.Checked.ToString());

            DateConfig["SingleEditor"] = helper.BaseElement.OuterXml;
            DateConfig.Save(); //�x�s���w�]��

            #endregion
        }

        private void InitializeRadioButton()
        {
            #region ���m���O�إ�
            DSResponse dsrsp = Config.GetAbsenceList();
            DSXmlHelper helper = dsrsp.GetContent();
            foreach (XmlElement element in helper.GetElements("Absence"))
            {
                AbsenceInfo info = new AbsenceInfo(element);

                //���䤣����
                if (!_absenceList.ContainsKey(info.Hotkey.ToUpper()))
                {
                    _absenceList.Add(info.Hotkey.ToUpper(), info);
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("���m�O�G{0}\n����G{1} �w����\n(�^��r���j�p�g�����ۦP����)");
                    MsgBox.Show(string.Format(sb.ToString(), info.Name, info.Hotkey));
                }

                RadioButton rb = new RadioButton();
                rb.Text = info.Name + "(" + info.Hotkey + ")";
                rb.AutoSize = true;
                //rb.Font = new Font(Framework.DotNetBar.FontStyles.GeneralFontFamily, 9.25f);
                rb.Tag = info;
                rb.CheckedChanged += new EventHandler(rb_CheckedChanged);
                rb.Click += new EventHandler(rb_CheckedChanged);
                panel.Controls.Add(rb);
                if (_checkedAbsence == null)
                {
                    _checkedAbsence = info;
                    rb.Checked = true;
                }
            }
            #endregion
        }

        void rb_CheckedChanged(object sender, EventArgs e)
        {
            #region ���m���O�إ�(�ƥ�)
            RadioButton rb = sender as RadioButton;
            if (rb.Checked)
            {
                _checkedAbsence = rb.Tag as AbsenceInfo;
                foreach (DataGridViewCell cell in dataGridView.SelectedCells)
                {
                    if (cell.ColumnIndex < _startIndex || cell.OwningRow.Visible == false) continue;
                    cell.Value = _checkedAbsence.Abbreviation;
                    AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
                    if (acInfo == null)
                    {
                        acInfo = new AbsenceCellInfo();
                    }
                    acInfo.SetValue(_checkedAbsence);
                    cell.Value = acInfo.AbsenceInfo.Abbreviation;
                    cell.Tag = acInfo;
                }
                dataGridView.Focus();
            }
            #endregion
        }

        private void InitializeDataGridViewColumn()
        {
            #region DataGridView��Column�إ�

            ColumnIndex.Clear(); //�M��

            DSResponse dsrsp = Config.GetPeriodList();
            DSXmlHelper helper = dsrsp.GetContent();
            PeriodCollection collection = new PeriodCollection();
            foreach (XmlElement element in helper.GetElements("Period"))
            {
                PeriodInfo info = new PeriodInfo(element);
                collection.Items.Add(info);
            }

            //SetColumn("colClassName", "�Z��", true);
            //SetColumn("colSeatNo", "�y��", true);
            //int ColumnsIndex = SetColumn("colStudentName", "�m�W", true);
            //dataGridView.Columns[ColumnsIndex].Frozen = true;
            //SetColumn("colStudentNumber", "�Ǹ�", true);
            //SetColumn("colDate", "���", true);
            //SetColumn("colWeek", "�P��", true);
            //SetColumn("colSchoolYear", "�Ǧ~��", false);
            //SetColumn("colSemester", "�Ǵ�", false);

            int ColumnsIndex = dataGridView.Columns.Add("colClassName", "�Z��");
            ColumnIndex.Add("�Z��", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colSeatNo", "�y��");
            ColumnIndex.Add("�y��", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colStudentName", "�m�W");
            ColumnIndex.Add("�m�W", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;
            dataGridView.Columns[ColumnsIndex].Frozen = true;

            ColumnsIndex = dataGridView.Columns.Add("colStudentNumber", "�Ǹ�");
            ColumnIndex.Add("�Ǹ�", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colDate", "���");
            ColumnIndex.Add("���", ColumnsIndex);
            //dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;
            dataGridView.Columns[ColumnsIndex].Width = 120;

            ColumnsIndex = dataGridView.Columns.Add("colWeek", "�P��");
            ColumnIndex.Add("�P��", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colSchoolYear", "�Ǧ~��");
            ColumnIndex.Add("�Ǧ~��", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = false;

            ColumnsIndex = dataGridView.Columns.Add("colSemester", "�Ǵ�");
            ColumnIndex.Add("�Ǵ�", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = false;

            _startIndex = ColumnIndex["�Ǵ�"] + 1;

            foreach (PeriodInfo info in collection.GetSortedList())
            {
                int columnIndex = dataGridView.Columns.Add(info.Name, info.Name);
                ColumnIndex.Add(info.Name, columnIndex); //�`��
                DataGridViewColumn column = dataGridView.Columns[columnIndex];
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                column.ReadOnly = true;
                column.Tag = info;
            }

            #endregion
        }

        //private int SetColumn(string HeaderName, string HeaderTitle, bool readOnly)
        //{
        //    int ColumnsIndex = dataGridView.Columns.Add(HeaderName, HeaderTitle);
        //    ColumnIndex.Add(HeaderTitle, ColumnsIndex);
        //    dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
        //    dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
        //    dataGridView.Columns[ColumnsIndex].ReadOnly = readOnly;
        //    dataGridView.Columns[ColumnsIndex].DefaultCellStyle.BackColor = Color.LightCyan;
        //    return ColumnsIndex;
        //}

        private int SortStudentInClass(SHStudentRecord xStud, SHStudentRecord yStud)
        {
            string xClass = xStud.Class != null ? xStud.Class.Name : "";
            xClass = xClass.PadLeft(6, '0');
            string yClass = yStud.Class != null ? yStud.Class.Name : "";
            yClass = yClass.PadLeft(6, '0');

            string xSean = xStud.SeatNo.HasValue ? xStud.SeatNo.Value.ToString() : "";
            xSean = xSean.PadLeft(6, '0');
            string ySean = yStud.SeatNo.HasValue ? yStud.SeatNo.Value.ToString() : "";
            ySean = ySean.PadLeft(6, '0');

            string xx = xClass + xSean;
            string yy = yClass + ySean;
            return xx.CompareTo(yy);
        }

        private void SearchDateRange()
        {
            #region ������
            DateTime start = dateTimeInput1.Value;
            DateTime end = dateTimeInput2.Value;

            dataGridView.Rows.Clear();

            TimeSpan ts = dateTimeInput2.Value - dateTimeInput1.Value;
            if (ts.Days > 1500)
            {
                FISCA.Presentation.Controls.MsgBox.Show("�z����F" + ts.Days.ToString() + "��\n�ѩ�������϶��L��,�Э��s�]�w����I");
                _currentStartDate = dateTimeInput1.Value = DateTime.Today;
                _currentEndDate = dateTimeInput2.Value = DateTime.Today;
                return;
            }

            List<SHStudentRecord> CatchStudentList = SHStudent.SelectByIDs(_studentList);

            CatchStudentList = SortClassIndex.SHSchoolData_SHStudentRecord(CatchStudentList);

            bool ColorTrue = true;
            foreach (SHStudentRecord each in CatchStudentList)
            {
                DateTime date = start;
                if (ColorTrue)
                {
                    ColorTrue = false;
                }
                else
                {
                    ColorTrue = true;
                }

                while (date.CompareTo(end) <= 0)
                {
                    if (!nowWeekDay.Contains(date.DayOfWeek)) //�o��
                    {
                        date = date.AddDays(1);
                        continue;
                    }
                    string dateValue = date.ToShortDateString();

                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dataGridView);
                    if (ColorTrue)
                    {
                        SetDataGridViewColor(row, ColorTrue);
                    }
                    else
                    {
                        SetDataGridViewColor(row, ColorTrue);
                    }
                    RowTag tag = new RowTag();
                    tag.Date = date;
                    tag.IsNew = true;
                    row.Tag = tag; //RowTag
                    row.Cells[0].Tag = each.ID; //�t�νs��

                    row.Cells[ColumnIndex["�Z��"]].Value = each.Class != null ? each.Class.Name : "";
                    row.Cells[ColumnIndex["�y��"]].Value = each.SeatNo.HasValue ? each.SeatNo.Value.ToString() : "";
                    row.Cells[ColumnIndex["�m�W"]].Value = each.Name;
                    row.Cells[ColumnIndex["�Ǹ�"]].Value = each.StudentNumber;

                    row.Cells[ColumnIndex["���"]].Value = dateValue;
                    row.Cells[ColumnIndex["�P��"]].Value = GetDayOfWeekInChinese(date.DayOfWeek);
                    _semesterProvider.SetDate(date);
                    row.Cells[ColumnIndex["�Ǧ~��"]].Value = _semesterProvider.SchoolYear.ToString();
                    row.Cells[ColumnIndex["�Ǵ�"]].Value = _semesterProvider.Semester.ToString();
                    date = date.AddDays(1);

                    dataGridView.Rows.Add(row);
                }
            }
            #endregion
        }

        private void SetDataGridViewColor(DataGridViewRow row,bool ColorMode)
        {
            foreach (DataGridViewCell cell in row.Cells)
            {
                if (ColorMode)
                {
                    cell.Style.BackColor = Color.LightCyan;
                }
                else
                {
                    cell.Style.BackColor = Color.White;
                }
            }
        }

        private void GetAbsense()
        {

            LOG.Clear();

            #region ���o���m�O��
            List<SHAttendanceRecord> attendList = new List<SHAttendanceRecord>();

            foreach (string each in _studentList)
            {
                if (!LOG.ContainsKey(each))
                {
                    LOG.Add(each, new LogStudent());
                }
            }

            foreach (SHAttendanceRecord each in SHAttendance.SelectByDate(dateTimeInput1.Value, dateTimeInput2.Value))
            {
                if (_studentList.Contains(each.RefStudentID))
                {
                    attendList.Add(each);
                }
            }

            foreach (SHAttendanceRecord attendanceRecord in attendList)
            {
                // �o�̭n���@�ǨƱ�  �Ҧp���F���i�h
                string occurDate = attendanceRecord.OccurDate.ToShortDateString();
                string schoolYear = attendanceRecord.SchoolYear.ToString();
                string semester = attendanceRecord.Semester.ToString();
                string id = attendanceRecord.ID;
                List<K12.Data.AttendancePeriod> dNode = attendanceRecord.PeriodDetail;

                //log �����ק�e����� �������
                DateTime logDate;
                if (DateTime.TryParse(occurDate, out logDate))
                {
                    if (!LOG[attendanceRecord.RefStudentID].beforeData.ContainsKey(logDate.ToShortDateString()))
                        LOG[attendanceRecord.RefStudentID].beforeData.Add(logDate.ToShortDateString(), new Dictionary<string, string>());
                }

                DataGridViewRow row = null;
                foreach (DataGridViewRow r in dataGridView.Rows)
                {
                    if (r.Cells[0].Tag as string == attendanceRecord.RefStudentID && "" + r.Cells[ColumnIndex["���"]].Value == attendanceRecord.OccurDate.ToShortDateString())
                    {
                        row = r;
                        break;
                    }
                }

                if (row == null) continue;
                RowTag rowTag = row.Tag as RowTag;
                rowTag.IsNew = false;
                rowTag.Key = id;

                row.Cells[0].Tag = attendanceRecord; //�����x�s��Cell[0]

                row.Cells[ColumnIndex["�Ǧ~��"]].Value = schoolYear;
                row.Cells[ColumnIndex["�Ǧ~��"]].Tag = new SemesterCellInfo(schoolYear);

                row.Cells[ColumnIndex["�Ǵ�"]].Value = semester;
                row.Cells[ColumnIndex["�Ǵ�"]].Tag = new SemesterCellInfo(semester);

                for (int i = _startIndex; i < dataGridView.Columns.Count; i++)
                {
                    DataGridViewColumn column = dataGridView.Columns[i];
                    PeriodInfo info = column.Tag as PeriodInfo;

                    foreach (K12.Data.AttendancePeriod node in dNode)
                    {
                        if (node.Period != info.Name) continue;
                        if (node.AbsenceType == null) continue;

                        DataGridViewCell cell = row.Cells[i];
                        foreach (AbsenceInfo ai in _absenceList.Values)
                        {
                            if (ai.Name != node.AbsenceType) continue;
                            AbsenceInfo ainfo = ai.Clone();
                            cell.Tag = new AbsenceCellInfo(ainfo);
                            cell.Value = ai.Abbreviation;

                            ////log �����ק�e����� ���m���ӳ���
                            if (!LOG[node.RefStudentID].beforeData[logDate.ToShortDateString()].ContainsKey(info.Name))
                                LOG[node.RefStudentID].beforeData[logDate.ToShortDateString()].Add(info.Name, ai.Name);

                            break;
                        }
                    }
                }
            } 
            #endregion
        }

        //�x�s
        private void btnSave_Click(object sender, EventArgs e)
        {
            #region Save
            if (!IsValid())
            {
                FISCA.Presentation.Controls.MsgBox.Show("������ҥ��ѡA�Эץ���A���x�s", "���ҥ���", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            List<SHAttendanceRecord> InsertHelper = new List<SHAttendanceRecord>(); //�s�W
            List<SHAttendanceRecord> updateHelper = new List<SHAttendanceRecord>(); //��s
            List<string> deleteList = new List<string>(); //�M��

            //List<string> synclist = new List<string>();

            ISemester semester = SemesterProvider.GetInstance();
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                RowTag tag = row.Tag as RowTag;
                semester.SetDate(tag.Date);

                ////log �����ק�᪺��� �������
                if (row.Cells[0].Tag is string)
                {
                    if (!LOG[row.Cells[0].Tag.ToString()].afterData.ContainsKey(tag.Date.ToShortDateString()))
                        LOG[row.Cells[0].Tag.ToString()].afterData.Add(tag.Date.ToShortDateString(), new Dictionary<string, string>());
                }
                else
                {
                    SHAttendanceRecord attRecord = row.Cells[0].Tag as SHAttendanceRecord;

                    if (!LOG[attRecord.RefStudentID].afterData.ContainsKey(tag.Date.ToShortDateString()))
                        LOG[attRecord.RefStudentID].afterData.Add(tag.Date.ToShortDateString(), new Dictionary<string, string>());
                }

                if (tag.IsNew)
                {
                    #region IsNew
                    string studentID = row.Cells[0].Tag as string;

                    SHAttendanceRecord attRecord = new SHAttendanceRecord();

                    bool hasContent = false;
                    for (int i = _startIndex; i < dataGridView.Columns.Count; i++)
                    {
                        DataGridViewCell cell = row.Cells[i];
                        if (string.IsNullOrEmpty(("" + cell.Value).Trim())) continue;

                        PeriodInfo pinfo = dataGridView.Columns[i].Tag as PeriodInfo;
                        AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
                        AbsenceInfo ainfo = acInfo.AbsenceInfo;

                        K12.Data.AttendancePeriod ap = new K12.Data.AttendancePeriod();
                        ap.Period = pinfo.Name;
                        ap.AbsenceType = ainfo.Name;
                        attRecord.PeriodDetail.Add(ap);

                        hasContent = true;

                        //log �����ק�᪺��� ���m���ӳ���
                        if (!LOG[studentID].afterData[tag.Date.ToShortDateString()].ContainsKey(pinfo.Name))
                            LOG[studentID].afterData[tag.Date.ToShortDateString()].Add(pinfo.Name, ainfo.Name);

                    }

                    if (hasContent)
                    {
                        attRecord.RefStudentID = studentID;
                        attRecord.SchoolYear = int.Parse("" + row.Cells[ColumnIndex["�Ǧ~��"]].Value);
                        attRecord.Semester = int.Parse("" + row.Cells[ColumnIndex["�Ǵ�"]].Value);
                        attRecord.OccurDate = DateTime.Parse(tag.Date.ToShortDateString());
                        InsertHelper.Add(attRecord);
                    }

                    #endregion
                }
                else // �Y�O�쥻�N��������
                {
                    #region �O�ª�

                    SHAttendanceRecord attRecord = row.Cells[0].Tag as SHAttendanceRecord;
                    attRecord.PeriodDetail.Clear(); //�M��

                    bool hasContent = false;
                    for (int i = _startIndex; i < dataGridView.Columns.Count; i++)
                    {
                        DataGridViewCell cell = row.Cells[i];
                        if (string.IsNullOrEmpty(("" + cell.Value).Trim())) continue;

                        PeriodInfo pinfo = dataGridView.Columns[i].Tag as PeriodInfo;
                        AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
                        AbsenceInfo ainfo = acInfo.AbsenceInfo;

                        K12.Data.AttendancePeriod ap = new K12.Data.AttendancePeriod();
                        ap.Period = pinfo.Name;
                        ap.AbsenceType = ainfo.Name;
                        attRecord.PeriodDetail.Add(ap);

                        hasContent = true;

                        //log �����ק�᪺��� ���m���ӳ���
                        if (!LOG[attRecord.RefStudentID].afterData[tag.Date.ToShortDateString()].ContainsKey(pinfo.Name))
                            LOG[attRecord.RefStudentID].afterData[tag.Date.ToShortDateString()].Add(pinfo.Name, ainfo.Name);
                    }

                    if (hasContent)
                    {
                        attRecord.SchoolYear = int.Parse("" + row.Cells[ColumnIndex["�Ǧ~��"]].Value);
                        attRecord.Semester = int.Parse("" + row.Cells[ColumnIndex["�Ǵ�"]].Value);
                        updateHelper.Add(attRecord);
                    }
                    else
                    {
                        deleteList.Add(tag.Key);

                        //log �����Q�R�������
                        LOG[attRecord.RefStudentID].afterData.Remove(tag.Date.ToShortDateString());
                        LOG[attRecord.RefStudentID].deleteData.Add(tag.Date.ToShortDateString());
                    }
                    #endregion
                }
            }

            #region InsertHelper
            if (InsertHelper.Count > 0)
            {
                try
                {
                    SHAttendance.Insert(InsertHelper);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���m�����s�W���� : " + ex.Message, "�s�W����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                //log �g�Jlog
                foreach (string each in LOG.Keys)
                {
                    foreach (string date in LOG[each].afterData.Keys)
                    {
                        if (!LOG[each].beforeData.ContainsKey(date) && LOG[each].afterData[date].Count > 0)
                        {
                            StringBuilder desc = new StringBuilder("");
                            desc.AppendLine("�ǥ͡u" + K12.Data.Student.SelectByID(each).Name + "�v");
                            desc.AppendLine("����u" + date + "�v");
                            foreach (string period in LOG[each].afterData[date].Keys)
                            {
                                desc.AppendLine("�`���u" + period + "�v�]���u" + LOG[each].afterData[date][period] + "�v");
                            }
                            ApplicationLog.Log("�ǰȨt��.���m���", "�妸�s�W���m���", "student", each, desc.ToString());
                            //Log����
                            //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Insert, _student.ID, desc.ToString(), this.Text, "");
                        }
                    }
                }
            }
            #endregion

            #region updateHelper
            if (updateHelper.Count > 0)
            {
                try
                {
                    SHAttendance.Update(updateHelper);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���m������s���� : " + ex.Message, "��s����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log �g�Jlog
                foreach (string each in LOG.Keys)
                {
                    foreach (string date in LOG[each].afterData.Keys)
                    {
                        if (LOG[each].beforeData.ContainsKey(date) && LOG[each].afterData[date].Count > 0)
                        {
                            bool dirty = false;
                            StringBuilder desc = new StringBuilder("");
                            desc.AppendLine("�ǥ͡u" + K12.Data.Student.SelectByID(each).Name + "�v ");
                            desc.AppendLine("����u" + date + "�v ");
                            foreach (string period in LOG[each].beforeData[date].Keys)
                            {
                                if (!LOG[each].afterData[date].ContainsKey(period))
                                    LOG[each].afterData[date].Add(period, "");
                            }
                            foreach (string period in LOG[each].afterData[date].Keys)
                            {
                                if (LOG[each].beforeData[date].ContainsKey(period))
                                {
                                    if (LOG[each].beforeData[date][period] != LOG[each].afterData[date][period])
                                    {
                                        dirty = true;
                                        desc.AppendLine("�`���u" + period + "�v�ѡu" + LOG[each].beforeData[date][period] + "�v�ܧ󬰡u" + LOG[each].afterData[date][period] + "�v");
                                    }
                                }
                                else
                                {
                                    dirty = true;
                                    desc.AppendLine("�`���u" + period + "�v�ѡu�v�ܧ󬰡u" + LOG[each].afterData[date][period] + "�v ");
                                }

                            }
                            if (dirty)
                            {
                                //Log����
                                ApplicationLog.Log("�ǰȨt��.���m���", "�妸�ק���m���", "student", each, desc.ToString());
                                //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Update, _student.ID, desc.ToString(), this.Text, "");
                            }
                        }
                    }
                }
            }
            #endregion

            #region deleteList
            if (deleteList.Count > 0)
            {

                try
                {
                    SHAttendance.Delete(deleteList);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���m�����R������ : " + ex.Message, "�R������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log �g�J�Q�R������ƪ�log
                foreach (string each in LOG.Keys)
                {
                    StringBuilder desc = new StringBuilder("");
                    desc.AppendLine("�ǥ͡u" + K12.Data.Student.SelectByID(each).Name + "�v");
                    foreach (string date in LOG[each].deleteData)
                    {
                        desc.AppendLine("�R���u" + date + "�v���m���� ");
                    }

                    //Log����
                    ApplicationLog.Log("�ǰȨt��.���m���", "�妸�R�����m���", "student", each, desc.ToString());
                    //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Delete, _student.ID, desc.ToString(), this.Text, "");
                }
            }
            #endregion

            //Ĳ�o�ܧ�ƥ�
            //Attendance.Instance.SyncDataBackground(_studentList);

            FISCA.Presentation.Controls.MsgBox.Show("�x�s���m��Ʀ��\!", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
            #endregion
        }

        private bool IsValid()
        {
            #region DataGridView�������(�p�GErrorText���e����)
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ErrorText != string.Empty)
                        return false;
                }
            }
            return true;
            #endregion
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            LoadAbsense(); //���s��z
        }

        private void LoadAbsense()
        {
            #region �B�z�P���]�w��
            WeekDay.Clear();
            nowWeekDay.Clear();
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["���m�妸�n��_�P���]�w_ByMany"];
            string cdIN = cd["�P���]�w"];

            XmlElement day;

            if (cdIN != "")
            {
                day = XmlHelper.LoadXml(cdIN);
            }
            else
            {
                day = null;
            }

            if (day != null)
            {
                foreach (XmlNode each in day.SelectNodes("Day"))
                {
                    XmlElement each2 = each as XmlElement;
                    WeekDay.Add(each2.GetAttribute("Detail"));
                }
            }
            else
            {
                WeekDay.AddRange(new string[] { "�P���@", "�P���G", "�P���T", "�P���|", "�P����", "�P����", "�P����" });
            }

            nowWeekDay = ChengDayOfWeel(WeekDay);
            #endregion

            SearchDateRange();
            GetAbsense();
            chkHasData_CheckedChanged(null, null);

        }

        private List<DayOfWeek> ChengDayOfWeel(List<string> list)
        {
            #region ���o�P����Ӫ�
            List<DayOfWeek> DOW = new List<DayOfWeek>();
            foreach (string each in list)
            {
                if (each == "�P���@")
                {
                    DOW.Add(DayOfWeek.Monday);
                }
                else if (each == "�P���G")
                {
                    DOW.Add(DayOfWeek.Tuesday);
                }
                else if (each == "�P���T")
                {
                    DOW.Add(DayOfWeek.Wednesday);
                }
                else if (each == "�P���|")
                {
                    DOW.Add(DayOfWeek.Thursday);
                }
                else if (each == "�P����")
                {
                    DOW.Add(DayOfWeek.Friday);
                }
                else if (each == "�P����")
                {
                    DOW.Add(DayOfWeek.Saturday);
                }
                else if (each == "�P����")
                {
                    DOW.Add(DayOfWeek.Sunday);
                }
            }

            return DOW; 
            #endregion
        }

        private string GetDayOfWeekInChinese(DayOfWeek day)
        {
            #region �P��(��/�^)��Ӫ�
            switch (day)
            {
                case DayOfWeek.Monday:
                    return "�@";
                case DayOfWeek.Tuesday:
                    return "�G";
                case DayOfWeek.Wednesday:
                    return "�T";
                case DayOfWeek.Thursday:
                    return "�|";
                case DayOfWeek.Friday:
                    return "��";
                case DayOfWeek.Saturday:
                    return "��";
                default:
                    return "��";
            } 
            #endregion
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("��Ƥw�ܧ�B�|���x�s�A�O�_���w�s����?", "�T�{", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    this.Close();
                }
            }
            else
                this.Close();
        }
        
        private void dataGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            #region CellDoubleClick
            if (e.Button != MouseButtons.Left) return;
            if (e.ColumnIndex < _startIndex) return;
            if (e.RowIndex < 0) return;
            DataGridViewCell cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
            cell.Value = _checkedAbsence.Abbreviation;
            AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
            if (acInfo == null)
            {
                acInfo = new AbsenceCellInfo();
            }
            acInfo.SetValue(_checkedAbsence);
            if (acInfo.IsValueChanged)
                cell.Value = acInfo.AbsenceInfo.Abbreviation;
            else
            {
                cell.Value = string.Empty;
                acInfo.SetValue(AbsenceInfo.Empty);
            }
            cell.Tag = acInfo; 
            #endregion
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            #region �Ǧ~��/�Ǵ���J����
            DataGridViewCell cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
            if (e.ColumnIndex == ColumnIndex["�Ǧ~��"])
            {
                string errorMessage = "";
                int schoolYear;
                if (cell.Value == null)
                    errorMessage = "�Ǧ~�פ��i���ť�";
                else if (!int.TryParse(cell.Value.ToString(), out schoolYear))
                    errorMessage = "�Ǧ~�ץ��������";

                if (errorMessage != "")
                {
                    cell.Style.BackColor = Color.Red;
                    cell.ToolTipText = errorMessage;
                }
                else
                {
                    cell.Style.BackColor = Color.White;
                    cell.ToolTipText = "";
                }
            }
            else if (e.ColumnIndex == ColumnIndex["�Ǵ�"])
            {
                string errorMessage = "";

                if (cell.Value == null)
                    errorMessage = "�Ǵ����i���ť�";
                else if (cell.Value.ToString() != "1" && cell.Value.ToString() != "2")
                    errorMessage = "�Ǵ���������ơy1�z�Ρy2�z";

                if (errorMessage != "")
                {
                    cell.ErrorText = errorMessage;
                }
                else
                {
                    cell.ErrorText = string.Empty;
                }
            } 
            #endregion
        }

        private void dataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            #region �p�G���U���s
            string key = KeyConverter.GetKeyMapping(e);

            if (!_absenceList.ContainsKey(key))
            {
                if (e.KeyCode != Keys.Space && e.KeyCode != Keys.Delete) return;
                foreach (DataGridViewCell cell in dataGridView.SelectedCells)
                {
                    if (cell.ColumnIndex < _startIndex || cell.OwningRow.Visible == false) continue;
                    cell.Value = null;
                    AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
                    if (acInfo != null)
                        acInfo.SetValue(null);
                }
            }
            else
            {
                AbsenceInfo info = _absenceList[key];
                foreach (DataGridViewCell cell in dataGridView.SelectedCells)
                {
                    if (cell.ColumnIndex < _startIndex || cell.OwningRow.Visible == false) continue;
                    AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;

                    if (acInfo == null)
                    {
                        acInfo = new AbsenceCellInfo();
                    }
                    acInfo.SetValue(info);

                    if (acInfo.IsValueChanged)
                        cell.Value = acInfo.AbsenceInfo.Abbreviation;
                    else
                    {
                        cell.Value = string.Empty;
                        acInfo.SetValue(AbsenceInfo.Empty);
                    }
                    cell.Tag = acInfo;
                }
            } 
            #endregion
        }
                    
        private void dateTimeInput1_Validated(object sender, EventArgs e)
        {
            #region dateTimeInput1����ܧ�ƥ�
            _errorProvider.SetError(dateTimeInput1, string.Empty);

            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("��Ƥw�ܧ�B�|���x�s�A�O�_���w�s����?", "�T�{", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    dateTimeInput1.Value = _currentStartDate;
                    return;
                }
            }
            _currentStartDate = dateTimeInput1.Value;
            dataGridView.Rows.Clear();
            LoadAbsense(); 
            #endregion
        }


        private void dateTimeInput2_Validated(object sender, EventArgs e)
        {
            #region dateTimeInput1����ܧ�ƥ�
            _errorProvider.SetError(dateTimeInput2, string.Empty);

            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("��Ƥw�ܧ�B�|���x�s�A�O�_���w�s����?", "�T�{", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    dateTimeInput2.Value = _currentEndDate;
                    return;
                }
            }
            _currentEndDate = dateTimeInput2.Value;
            dataGridView.Rows.Clear();
            LoadAbsense();  
            #endregion  
        } 

        private bool IsDirty()
        {
            #region �������
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Tag == null) continue;
                    if (cell.Tag is SemesterCellInfo)
                    {
                        SemesterCellInfo cInfo = cell.Tag as SemesterCellInfo;
                        if (cInfo.IsDirty) return true;
                    }
                    else if (cell.Tag is AbsenceCellInfo)
                    {
                        AbsenceCellInfo cInfo = cell.Tag as AbsenceCellInfo;
                        if (cInfo.IsDirty) return true;
                    }
                }
            }
            return false; 
            #endregion
        }

        private void chkHasData_CheckedChanged(object sender, EventArgs e)
        {
            #region ����ܦ����m�����
            dataGridView.SuspendLayout();

            if (chkHasData.Checked == true)
            {
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    bool hasData = false;
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.ColumnIndex < _startIndex || cell.OwningRow.Visible == false) continue;
                        if (!string.IsNullOrEmpty("" + cell.Value))
                        {
                            hasData = true;
                            break;
                        }
                    }
                    if (hasData == false)
                    {
                        _hiddenRows.Add(row);
                        row.Visible = false;
                    }
                }
            }
            else
            {
                foreach (DataGridViewRow row in _hiddenRows)
                    row.Visible = true;
            }

            dataGridView.ResumeLayout(); 
            #endregion
        }

        private void btnDay_Click(object sender, EventArgs e)
        {
            #region �P���]�w
            Searchday Sday = new Searchday("���m�妸�n��_�P���]�w_ByMany");
            if (Sday.ShowDialog() == DialogResult.Yes)
            {
                LoadAbsense();
            } 
            #endregion
        }

        private void dateTimeInput1_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView.Rows.Count != 0)
            {
                LoadAbsense();
            }
        }

        private void dateTimeInput2_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView.Rows.Count != 0)
            {
                LoadAbsense();
            }
        }

        private void SingleEditor_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveDateSetting();
        }

        private void dataGridView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                row.Cells[e.ColumnIndex].Selected = true;
            }
        }

    }
}