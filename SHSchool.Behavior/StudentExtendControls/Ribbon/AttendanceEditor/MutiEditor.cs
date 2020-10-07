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

        //log 需要用到的
        //<學生,<日期<節次,缺曠別>>
        //<學生ID,每一個Log記錄>
        private Dictionary<string, LogStudent> LOG = new Dictionary<string, LogStudent>();

        public MutiEditor(List<StudentRecord> students)
        {
            InitializeComponent(); //設計工具產生的

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
            #region 日期定義
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
                DateConfig.Save(); //儲存此預設檔
            }

            XmlElement loadXml = DSXmlHelper.LoadXml(date);
            checkBoxX1.Checked = bool.Parse(loadXml.SelectSingleNode("Locked").InnerText);

            if (checkBoxX1.Checked) //如果是鎖定,就取鎖定日期
            {
                dateTimeInput1.Text = loadXml.SelectSingleNode("Date").InnerText;
            }
            else //如果沒有鎖定,就取當天
            {
                dateTimeInput1.Text = DateTime.Today.ToShortDateString();
            }
            _currentDate = dateTimeInput1.Value;
            #endregion
        }

        private void SaveDateSetting()
        {
            #region 儲存日期資料
            K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["Attendance_BatchEditor"];

            DSXmlHelper helper = new DSXmlHelper("xml");
            helper.AddElement("Date");
            helper.AddText("Date", dateTimeInput1.Value.ToShortDateString());
            helper.AddElement("Locked");
            helper.AddText("Locked", checkBoxX1.Checked.ToString());

            DateConfig["MutiEditor"] = helper.BaseElement.OuterXml;
            DateConfig.Save(); //儲存此預設檔

            #endregion
        }

        private void InitializeRadioButton()
        {
            #region 缺曠類別建立
            DSResponse dsrsp = Framework.Feature.Config.GetAbsenceList();
            DSXmlHelper helper = dsrsp.GetContent();
            foreach (XmlElement element in helper.GetElements("Absence"))
            {
                AbsenceInfo info = new AbsenceInfo(element);
                //熱鍵不重覆
                if (!_absenceList.ContainsKey(info.Hotkey.ToUpper()))
                {
                    _absenceList.Add(info.Hotkey.ToUpper(), info);
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("缺曠別：{0}\n熱鍵：{1} 已重覆\n(英文字母大小寫視為相同熱鍵)");
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
            #region 缺曠類別建立(事件)
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
            #region DataGridView的Column建立

            ColumnIndex.Clear();

            DSResponse dsrsp = Framework.Feature.Config.GetPeriodList();
            DSXmlHelper helper = dsrsp.GetContent();
            PeriodCollection collection = new PeriodCollection();
            foreach (XmlElement element in helper.GetElements("Period"))
            {
                PeriodInfo info = new PeriodInfo(element);
                collection.Items.Add(info);
            }
            int ColumnsIndex = dataGridView.Columns.Add("colClassName", "班級");
            ColumnIndex.Add("班級", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colSeatNo", "座號");
            ColumnIndex.Add("座號", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colName", "姓名");
            ColumnIndex.Add("姓名", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;
            dataGridView.Columns[ColumnsIndex].Frozen = true; //由此開始可拖移

            ColumnsIndex = dataGridView.Columns.Add("colSchoolNumber", "學號");
            ColumnIndex.Add("學號", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colSchoolYear", "學年度");
            ColumnIndex.Add("學年度", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = false;

            ColumnsIndex = dataGridView.Columns.Add("colSemester", "學期");
            ColumnIndex.Add("學期", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = false;

            _startIndex = ColumnIndex["學期"] + 1;

            List<string> cols = new List<string>() { "學年度", "學期" };

            foreach (PeriodInfo info in collection.GetSortedList())
            {
                cols.Add(info.Name);

                int columnIndex = dataGridView.Columns.Add(info.Name, info.Name);
                ColumnIndex.Add(info.Name, columnIndex); //節次
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
            #region 日期選擇

            dataGridView.Rows.Clear();
            _semesterProvider.SetDate(dateTimeInput1.Value);
            _students = SortClassIndex.K12Data_StudentRecord(_students);

            foreach (StudentRecord student in _students)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView);
                row.Cells[0].Tag = student.ID;

                row.Cells[ColumnIndex["班級"]].Value = (student.Class != null) ? student.Class.Name : "";
                row.Cells[ColumnIndex["姓名"]].Value = student.Name;
                row.Cells[ColumnIndex["學號"]].Value = student.StudentNumber;
                row.Cells[ColumnIndex["座號"]].Value = student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "";

                row.Cells[ColumnIndex["學年度"]].Value = _semesterProvider.SchoolYear;
                row.Cells[ColumnIndex["學期"]].Value = _semesterProvider.Semester;
                row.Cells[ColumnIndex["學年度"]].Tag = new SemesterCellInfo(_semesterProvider.SchoolYear.ToString());
                row.Cells[ColumnIndex["學期"]].Tag = new SemesterCellInfo(_semesterProvider.Semester.ToString());
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
                // 這裡要做一些事情  例如找到東西塞進去
                string occurDate = element.OccurDate.ToShortDateString();
                string schoolYear = element.SchoolYear.ToString();
                string semester = element.Semester.ToString();
                string id = element.ID;
                List<K12.Data.AttendancePeriod> dNode = element.PeriodDetail;

                //log 紀錄修改前的資料 日期部分
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

                row.Cells[0].Tag = element; //把資料儲存於Cell[0]

                row.Cells[ColumnIndex["學年度"]].Value = schoolYear;
                row.Cells[ColumnIndex["學年度"]].Tag = new SemesterCellInfo(schoolYear);

                row.Cells[ColumnIndex["學期"]].Value = semester;
                row.Cells[ColumnIndex["學期"]].Tag = new SemesterCellInfo(semester);

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

                            ////log 紀錄修改前的資料 缺曠明細部分
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
                FISCA.Presentation.Controls.MsgBox.Show("資料驗證失敗，請修正後再行儲存", "驗證失敗", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            List<AttendanceRecord> InsertHelper = new List<AttendanceRecord>(); //新增
            List<AttendanceRecord> updateHelper = new List<AttendanceRecord>(); //更新
            List<string> deleteList = new List<string>(); //清空

            ISemester semester = SemesterProvider.GetInstance();
            foreach (DataGridViewRow row in dataGridView.Rows)
            {



                RowTag tag = row.Tag as RowTag;
                semester.SetDate(tag.Date);

                ////log 紀錄修改後的資料 日期部分
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

                        //log 紀錄修改後的資料 缺曠明細部分
                        if (!LOG[studentID].afterData[tag.Date.ToShortDateString()].ContainsKey(pinfo.Name))
                            LOG[studentID].afterData[tag.Date.ToShortDateString()].Add(pinfo.Name, ainfo.Name);

                    }

                    if (hasContent)
                    {
                        attRecord.RefStudentID = studentID;
                        attRecord.SchoolYear = int.Parse("" + row.Cells[ColumnIndex["學年度"]].Value);
                        attRecord.Semester = int.Parse("" + row.Cells[ColumnIndex["學期"]].Value);
                        attRecord.OccurDate = DateTime.Parse(tag.Date.ToShortDateString());
                        InsertHelper.Add(attRecord);
                    }

                    #endregion
                }
                else // 若是原本就有紀錄的
                {
                    #region 是舊的

                    AttendanceRecord attRecord = row.Cells[0].Tag as AttendanceRecord;
                    attRecord.PeriodDetail.Clear(); //清空

                    bool hasContent = false;
                    for (int i = _startIndex; i < dataGridView.Columns.Count; i++)
                    {
                        DataGridViewCell cell = row.Cells[i];
                        if (string.IsNullOrEmpty(("" + cell.Value).Trim())) continue; //避免是空白

                        PeriodInfo pinfo = dataGridView.Columns[i].Tag as PeriodInfo;
                        AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
                        AbsenceInfo ainfo = acInfo.AbsenceInfo;

                        K12.Data.AttendancePeriod ap = new K12.Data.AttendancePeriod();
                        ap.Period = pinfo.Name;
                        ap.AbsenceType = ainfo.Name;
                        attRecord.PeriodDetail.Add(ap);

                        hasContent = true;

                        //log 紀錄修改後的資料 缺曠明細部分
                        if (!LOG[attRecord.RefStudentID].afterData[tag.Date.ToShortDateString()].ContainsKey(pinfo.Name))
                            LOG[attRecord.RefStudentID].afterData[tag.Date.ToShortDateString()].Add(pinfo.Name, ainfo.Name);
                    }

                    if (hasContent)
                    {
                        attRecord.SchoolYear = int.Parse("" + row.Cells[ColumnIndex["學年度"]].Value);
                        attRecord.Semester = int.Parse("" + row.Cells[ColumnIndex["學期"]].Value);
                        updateHelper.Add(attRecord);
                    }
                    else
                    {
                        deleteList.Add(tag.Key);

                        //log 紀錄被刪除的資料
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
                    FISCA.Presentation.Controls.MsgBox.Show("缺曠紀錄新增失敗 : " + ex.Message, "新增失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                //log 寫入log
                foreach (string each in LOG.Keys)
                {
                    foreach (string date in LOG[each].afterData.Keys)
                    {
                        if (!LOG[each].beforeData.ContainsKey(date) && LOG[each].afterData[date].Count > 0)
                        {
                            StringBuilder desc = new StringBuilder("");
                            desc.AppendLine("學生「" + K12.Data.Student.SelectByID(each).Name + "」");
                            desc.AppendLine("日期「" + date + "」");
                            foreach (string period in LOG[each].afterData[date].Keys)
                            {
                                desc.AppendLine("節次「" + period + "」設為「" + LOG[each].afterData[date][period] + "」");
                            }
                            ApplicationLog.Log("學務系統.缺曠資料", "批次新增缺曠資料", "student", each, desc.ToString());
                            //Log部份
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
                    FISCA.Presentation.Controls.MsgBox.Show("缺曠紀錄更新失敗 : " + ex.Message, "更新失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log 寫入log
                foreach (string each in LOG.Keys)
                {
                    foreach (string date in LOG[each].afterData.Keys)
                    {
                        if (LOG[each].beforeData.ContainsKey(date) && LOG[each].afterData[date].Count > 0)
                        {
                            bool dirty = false;
                            StringBuilder desc = new StringBuilder("");
                            desc.AppendLine("學生「" + K12.Data.Student.SelectByID(each).Name + "」 ");
                            desc.AppendLine("日期「" + date + "」 ");
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
                                        desc.AppendLine("節次「" + period + "」由「" + LOG[each].beforeData[date][period] + "」變更為「" + LOG[each].afterData[date][period] + "」");
                                    }
                                }
                                else
                                {
                                    dirty = true;
                                    desc.AppendLine("節次「" + period + "」由「」變更為「" + LOG[each].afterData[date][period] + "」 ");
                                }

                            }
                            if (dirty)
                            {
                                //Log部份
                                ApplicationLog.Log("學務系統.缺曠資料", "批次修改缺曠資料", "student", each, desc.ToString());
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
                    FISCA.Presentation.Controls.MsgBox.Show("缺曠紀錄刪除失敗 : " + ex.Message, "刪除失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log 寫入被刪除的資料的log
                foreach (string each in LOG.Keys)
                {
                    StringBuilder desc = new StringBuilder("");
                    desc.AppendLine("學生「" + K12.Data.Student.SelectByID(each).Name + "」");
                    foreach (string date in LOG[each].deleteData)
                    {
                        desc.AppendLine("刪除「" + date + "」缺曠紀錄 ");
                    }

                    //Log部份
                    ApplicationLog.Log("學務系統.缺曠資料", "批次刪除缺曠資料", "student", each, desc.ToString());
                    //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Delete, _student.ID, desc.ToString(), this.Text, "");
                }
            }
            #endregion

            FISCA.Presentation.Controls.MsgBox.Show("儲存缺曠資料成功!", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();

            SaveDateSetting();       //儲存日期是否鎖定的設定
        }

        private bool IsValid()
        {
            #region DataGridView資料驗證(如果ErrorText內容為空)

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
            //    _errProvider.SetError(startDate, "日期格式錯誤");
            //    return;
            //}
            SearchStudentRange();
            GetAbsense();
            chkHasData_CheckedChanged(null, null);
        }

        private void dataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            #region 如果按下按鈕
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
            #region 學年度/學期輸入驗證
            DataGridViewCell cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
            DataGridViewColumn column = dataGridView.Columns[e.ColumnIndex];
            if (column.Index == ColumnIndex["學年度"])
            {
                string errorMessage = "";
                int schoolYear;
                if (cell.Value == null)
                    errorMessage = "學年度不可為空白";
                else if (!int.TryParse(cell.Value.ToString(), out schoolYear))
                    errorMessage = "學年度必須為整數";

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
            else if (column.Index == ColumnIndex["學期"])
            {
                string errorMessage = string.Empty;

                if (cell.Value == null)
                    errorMessage = "學期不可為空白";
                else if (cell.Value.ToString() != "1" && cell.Value.ToString() != "2")
                    errorMessage = "學期必須為整數『1』或『2』";

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
                if (FISCA.Presentation.Controls.MsgBox.Show("資料已變更且尚未儲存，是否放棄已編輯資料?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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
                if (FISCA.Presentation.Controls.MsgBox.Show("資料已變更且尚未儲存，是否放棄已編輯資料?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
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
            //    toolTip.SetToolTip(picLock, "缺曠日期為未鎖定狀態，您可以點選圖示，將特定日期鎖定。");
            //    labelX2.Text = "";
            //}
            //else
            //{
            //    picLock.Image = Resources._lock;
            //    picLock.Tag = true;
            //    toolTip.SetToolTip(picLock, "缺曠日期已鎖定，您可以點選圖示解除鎖定。");
            //    labelX2.Text = "已鎖定缺曠日期";
            //}

            //this.SaveDateSetting();
        }

        private bool IsDirty()
        { 
            #region 資料驗證
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Tag == null) continue; //Tag是空的
                    if (cell.Tag is SemesterCellInfo) //學年度學期
                    {
                        SemesterCellInfo cInfo = cell.Tag as SemesterCellInfo;
                        if (cInfo.IsDirty) return true;
                    }
                    else if (cell.Tag is AbsenceCellInfo) //缺曠別
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
            #region 僅顯示有缺曠的資料
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

        //關閉視窗即儲存設定
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