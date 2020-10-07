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
using K12.Data;
using SHSchool.Behavior.Feature;
using FISCA.Presentation.Controls;

namespace SHSchool.Behavior.StudentExtendControls
{
    public partial class MutiEditor : FISCA.Presentation.Controls.BaseForm
    {
        private List<StudentRecord> _students;
        private ISemester _semesterProvider;
        private Dictionary<string, AbsenceInfo> _absenceList;
        private int _startIndex;
        private AbsenceInfo _checkedAbsence;
        private DateTime _currentDate;
        List<DataGridViewRow> _hiddenRows;

        Dictionary<string, int> ColumnIndex = new Dictionary<string, int>();

        //System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();

        //log �ݭn�Ψ쪺
        //<�ǥ�,<���<�`��,���m�O>>
        //<�ǥ�ID,�C�@��Log�O��>
        private Dictionary<string, LogStudent> LOG = new Dictionary<string, LogStudent>();

        public MutiEditor(List<StudentRecord> students)
        {
            InitializeComponent(); //�]�p�u�㲣�ͪ�

            _students = students;
            _absenceList = new Dictionary<string, AbsenceInfo>();
            _semesterProvider = SemesterProvider.GetInstance();

            _hiddenRows = new List<DataGridViewRow>();
        }

        private void MutiEditor_Load(object sender, EventArgs e)
        {
            #region Load
            InitializeRadioButton();
            InitializeDateRange();
            InitializeDataGridView();
            //SearchStudentRange();
            btnRenew_Click(null, null);
            #endregion
        }

        private void InitializeDateRange()
        {
            #region ����w�q
            K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["Attendance_BatchEditor"];

            string date = DateConfig["MutiEditor"];

            if (date == "")
            {
                DSXmlHelper helper = new DSXmlHelper("xml");
                helper.AddElement("Date");
                helper.AddText("Date", DateTime.Today.ToShortDateString());
                helper.AddElement("Locked");
                helper.AddText("Locked", "false");

                date = helper.BaseElement.OuterXml;
                DateConfig["MutiEditor"] = date;
                DateConfig.Save(); //�x�s���w�]��
            }

            XmlElement loadXml = DSXmlHelper.LoadXml(date);
            checkBoxX1.Checked = bool.Parse(loadXml.SelectSingleNode("Locked").InnerText);

            if (checkBoxX1.Checked) //�p�G�O��w,�N����w���
            {
                dateTimeInput1.Text = loadXml.SelectSingleNode("Date").InnerText;
            }
            else //�p�G�S����w,�N�����
            {
                dateTimeInput1.Text = DateTime.Today.ToShortDateString();
            }
            _currentDate = dateTimeInput1.Value;
            #endregion
        }

        private void SaveDateSetting()
        {
            #region �x�s������
            K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["Attendance_BatchEditor"];

            DSXmlHelper helper = new DSXmlHelper("xml");
            helper.AddElement("Date");
            helper.AddText("Date", dateTimeInput1.Value.ToShortDateString());
            helper.AddElement("Locked");
            helper.AddText("Locked", checkBoxX1.Checked.ToString());

            DateConfig["MutiEditor"] = helper.BaseElement.OuterXml;
            DateConfig.Save(); //�x�s���w�]��

            #endregion
        }

        private void InitializeRadioButton()
        {
            #region ���m���O�إ�
            DSResponse dsrsp = Framework.Feature.Config.GetAbsenceList();
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

        private void InitializeDataGridView()
        {
            InitializeDataGridViewColumn();
        }

        private void InitializeDataGridViewColumn()
        {
            #region DataGridView��Column�إ�

            ColumnIndex.Clear();

            DSResponse dsrsp = Framework.Feature.Config.GetPeriodList();
            DSXmlHelper helper = dsrsp.GetContent();
            PeriodCollection collection = new PeriodCollection();
            foreach (XmlElement element in helper.GetElements("Period"))
            {
                PeriodInfo info = new PeriodInfo(element);
                collection.Items.Add(info);
            }
            int ColumnsIndex = dataGridView.Columns.Add("colClassName", "�Z��");
            ColumnIndex.Add("�Z��", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colSeatNo", "�y��");
            ColumnIndex.Add("�y��", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colName", "�m�W");
            ColumnIndex.Add("�m�W", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;
            dataGridView.Columns[ColumnsIndex].Frozen = true; //�Ѧ��}�l�i�첾

            ColumnsIndex = dataGridView.Columns.Add("colSchoolNumber", "�Ǹ�");
            ColumnIndex.Add("�Ǹ�", ColumnsIndex);
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

            List<string> cols = new List<string>() { "�Ǧ~��", "�Ǵ�" };

            foreach (PeriodInfo info in collection.GetSortedList())
            {
                cols.Add(info.Name);

                int columnIndex = dataGridView.Columns.Add(info.Name, info.Name);
                ColumnIndex.Add(info.Name, columnIndex); //�`��
                DataGridViewColumn column = dataGridView.Columns[columnIndex];
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
                column.ReadOnly = true;
                column.Tag = info;
            }

            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dataGridView, cols);


            #endregion
        }

        //private int SortStudent(StudentRecord sr1, StudentRecord sr2)
        //{
        //    return sr1.Name.CompareTo(sr2.Name);
        //}

        public int SortStudent(K12.Data.StudentRecord x, K12.Data.StudentRecord y)
        {
            K12.Data.StudentRecord student1 = x;
            K12.Data.StudentRecord student2 = y;

            string ClassName1 = student1.Class != null ? student1.Class.Name : "";
            ClassName1 = ClassName1.PadLeft(5, '0');
            string ClassName2 = student2.Class != null ? student2.Class.Name : "";
            ClassName2 = ClassName2.PadLeft(5, '0');

            string Sean1 = student1.SeatNo.HasValue ? student1.SeatNo.Value.ToString() : "";
            Sean1 = Sean1.PadLeft(3, '0');
            string Sean2 = student2.SeatNo.HasValue ? student2.SeatNo.Value.ToString() : "";
            Sean2 = Sean2.PadLeft(3, '0');

            ClassName1 += Sean1;
            ClassName2 += Sean2;

            return ClassName1.CompareTo(ClassName2);
        }

        private void SearchStudentRange()
        {
            #region ������

            dataGridView.Rows.Clear();
            _semesterProvider.SetDate(dateTimeInput1.Value);
            _students = SortClassIndex.K12Data_StudentRecord(_students);

            foreach (StudentRecord student in _students)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView);
                row.Cells[0].Tag = student.ID;

                row.Cells[ColumnIndex["�Z��"]].Value = (student.Class != null) ? student.Class.Name : "";
                row.Cells[ColumnIndex["�m�W"]].Value = student.Name;
                row.Cells[ColumnIndex["�Ǹ�"]].Value = student.StudentNumber;
                row.Cells[ColumnIndex["�y��"]].Value = student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "";

                row.Cells[ColumnIndex["�Ǧ~��"]].Value = _semesterProvider.SchoolYear;
                row.Cells[ColumnIndex["�Ǵ�"]].Value = _semesterProvider.Semester;
                row.Cells[ColumnIndex["�Ǧ~��"]].Tag = new SemesterCellInfo(_semesterProvider.SchoolYear.ToString());
                row.Cells[ColumnIndex["�Ǵ�"]].Tag = new SemesterCellInfo(_semesterProvider.Semester.ToString());
                RowTag tag = new RowTag();
                tag.Date = dateTimeInput1.Value;
                tag.IsNew = true;
                row.Tag = tag;

                dataGridView.Rows.Add(row);
            } 
            #endregion
        }

        private void GetAbsense()
        {
            LOG.Clear();

            List<AttendanceRecord> attendList = new List<AttendanceRecord>();
            List<string> list = new List<string>();
            foreach (StudentRecord each in _students)
            {
                if (!list.Contains(each.ID))
                {
                    list.Add(each.ID);

                    if (!LOG.ContainsKey(each.ID))
                    {
                        LOG.Add(each.ID, new LogStudent());
                    }
                }
            }
            foreach (AttendanceRecord each in Attendance.SelectByStudentIDs(list))
            {
                if (each.OccurDate.CompareTo(dateTimeInput1.Value) == 0)
                {
                    attendList.Add(each);
                }
            }

            foreach (AttendanceRecord element in attendList)
            {
                // �o�̭n���@�ǨƱ�  �Ҧp���F���i�h
                string occurDate = element.OccurDate.ToShortDateString();
                string schoolYear = element.SchoolYear.ToString();
                string semester = element.Semester.ToString();
                string id = element.ID;
                List<K12.Data.AttendancePeriod> dNode = element.PeriodDetail;

                //log �����ק�e����� �������
                DateTime logDate;
                if (DateTime.TryParse(occurDate, out logDate))
                {
                    if (!LOG[element.RefStudentID].beforeData.ContainsKey(logDate.ToShortDateString()))
                        LOG[element.RefStudentID].beforeData.Add(logDate.ToShortDateString(), new Dictionary<string, string>());
                }

                DataGridViewRow row = null;
                foreach (DataGridViewRow r in dataGridView.Rows)
                {
                    if (r.Cells[0].Tag as string == element.RefStudentID) //&& "" + r.Cells[1].Value == element.OccurDate.ToShortDateString())
                    {
                        row = r;
                        break;
                    }
                }

                if (row == null) continue;
                RowTag rowTag = row.Tag as RowTag;
                rowTag.IsNew = false;
                rowTag.Key = id;

                row.Cells[0].Tag = element; //�����x�s��Cell[0]

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
        }

        private void btnSave_Click(object sender, EventArgs e)
        {

            if (!IsValid())
            {
                FISCA.Presentation.Controls.MsgBox.Show("������ҥ��ѡA�Эץ���A���x�s", "���ҥ���", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            List<AttendanceRecord> InsertHelper = new List<AttendanceRecord>(); //�s�W
            List<AttendanceRecord> updateHelper = new List<AttendanceRecord>(); //��s
            List<string> deleteList = new List<string>(); //�M��

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
                    AttendanceRecord attRecord = row.Cells[0].Tag as AttendanceRecord;

                    if (!LOG[attRecord.RefStudentID].afterData.ContainsKey(tag.Date.ToShortDateString()))
                        LOG[attRecord.RefStudentID].afterData.Add(tag.Date.ToShortDateString(), new Dictionary<string, string>());
                }

                if (tag.IsNew)
                {
                    #region IsNew
                    string studentID = row.Cells[0].Tag as string;

                    AttendanceRecord attRecord = new AttendanceRecord();

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

                    AttendanceRecord attRecord = row.Cells[0].Tag as AttendanceRecord;
                    attRecord.PeriodDetail.Clear(); //�M��

                    bool hasContent = false;
                    for (int i = _startIndex; i < dataGridView.Columns.Count; i++)
                    {
                        DataGridViewCell cell = row.Cells[i];
                        if (string.IsNullOrEmpty(("" + cell.Value).Trim())) continue; //�קK�O�ť�

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
                    Attendance.Insert(InsertHelper);
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
                    Attendance.Update(updateHelper);
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
                    Attendance.Delete(deleteList);
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

            FISCA.Presentation.Controls.MsgBox.Show("�x�s���m��Ʀ��\!", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();

            SaveDateSetting();       //�x�s����O�_��w���]�w
        }

        private bool IsValid()
        {
            #region DataGridView�������(�p�GErrorText���e����)

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.ErrorText))
                        return false;
                }
            }
            return true; 
            #endregion
        }

        private void btnRenew_Click(object sender, EventArgs e)
        {
            //if (!startDate.IsValid)
            //{
            //    _errProvider.SetError(startDate, "����榡���~");
            //    return;
            //}
            SearchStudentRange();
            GetAbsense();
            chkHasData_CheckedChanged(null, null);
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
            DataGridViewColumn column = dataGridView.Columns[e.ColumnIndex];
            if (column.Index == ColumnIndex["�Ǧ~��"])
            {
                string errorMessage = "";
                int schoolYear;
                if (cell.Value == null)
                    errorMessage = "�Ǧ~�פ��i���ť�";
                else if (!int.TryParse(cell.Value.ToString(), out schoolYear))
                    errorMessage = "�Ǧ~�ץ��������";

                if (errorMessage != "")
                {
                    cell.ErrorText = errorMessage;
                    //cell.Style.BackColor = Color.Red;
                    //cell.ToolTipText = errorMessage;
                }
                else
                {
                    cell.ErrorText = string.Empty;
                    SemesterCellInfo cinfo = cell.Tag as SemesterCellInfo;
                    cinfo.SetValue(cell.Value == null ? string.Empty : cell.Value.ToString());
                    //cell.Style.BackColor = Color.White;
                    //cell.ToolTipText = "";
                }
            }
            else if (column.Index == ColumnIndex["�Ǵ�"])
            {
                string errorMessage = string.Empty;

                if (cell.Value == null)
                    errorMessage = "�Ǵ����i���ť�";
                else if (cell.Value.ToString() != "1" && cell.Value.ToString() != "2")
                    errorMessage = "�Ǵ���������ơy1�z�Ρy2�z";

                if (errorMessage != string.Empty)
                {
                    cell.ErrorText = errorMessage;
                    //cell.Style.BackColor = Color.Red;
                    //cell.ToolTipText = errorMessage;
                }
                else
                {
                    cell.ErrorText = string.Empty;
                    SemesterCellInfo cinfo = cell.Tag as SemesterCellInfo;
                    cinfo.SetValue(cell.Value == null ? string.Empty : cell.Value.ToString());
                    //cell.Style.BackColor = Color.White;
                    //cell.ToolTipText = "";
                }
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

        private void dateTimeInput1_Validated(object sender, EventArgs e)
        {
            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("��Ƥw�ܧ�B�|���x�s�A�O�_���w�s����?", "�T�{", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    dateTimeInput1.Value = _currentDate;
                    return;
                }
            }
            _currentDate = dateTimeInput1.Value;
            SearchStudentRange();
            GetAbsense();
            chkHasData_CheckedChanged(null, null);
        }
        
        private void picLock_Click(object sender, EventArgs e)
        {
            //bool isLock = false;
            //if (picLock.Tag != null)
            //{
            //    if (!bool.TryParse(picLock.Tag.ToString(), out isLock))
            //        isLock = false;
            //}
            //if (isLock)
            //{
            //    picLock.Image = Resources.unlock;
            //    picLock.Tag = false;
            //    toolTip.SetToolTip(picLock, "���m���������w���A�A�z�i�H�I��ϥܡA�N�S�w�����w�C");
            //    labelX2.Text = "";
            //}
            //else
            //{
            //    picLock.Image = Resources._lock;
            //    picLock.Tag = true;
            //    toolTip.SetToolTip(picLock, "���m����w��w�A�z�i�H�I��ϥܸѰ���w�C");
            //    labelX2.Text = "�w��w���m���";
            //}

            //this.SaveDateSetting();
        }

        private bool IsDirty()
        { 
            #region �������
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Tag == null) continue; //Tag�O�Ū�
                    if (cell.Tag is SemesterCellInfo) //�Ǧ~�׾Ǵ�
                    {
                        SemesterCellInfo cInfo = cell.Tag as SemesterCellInfo;
                        if (cInfo.IsDirty) return true;
                    }
                    else if (cell.Tag is AbsenceCellInfo) //���m�O
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

        private void dateTimeInput1_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView.Rows.Count != 0)
            {
                btnRenew_Click(null, null);
            }
        }

        //���������Y�x�s�]�w
        private void MutiEditor_FormClosing(object sender, FormClosingEventArgs e)
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

    //public class StudentRowTag
    //{
    //    private StudentRecord _student;
    //    public StudentRecord Student
    //    {
    //        get { return _student; }
    //        set { _student = value; }
    //    }
    //    private string _RowID;
    //    public string RowID
    //    {
    //        get { return _RowID; }
    //        set { _RowID = value; }
    //    }
    //}
}