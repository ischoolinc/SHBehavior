using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using Aspose.Cells;
using FISCA.DSAUtil;
using System.Xml;
using System.IO;
using K12.Data;
using SHSchool.Data;
using SmartSchool.ePaper;

namespace 德行成績試算表
{
    public partial class NewForm : BaseForm
    {
        /// <summary>
        /// 班級電子報表
        /// </summary>
        SmartSchool.ePaper.ElectronicPaper paperForClass { get; set; }
        bool Carty_paper = false;

        int _Schoolyear = 90;
        int _Semester = 1;
        int SizeIndex = 0;
        Dictionary<string, List<string>> Type = null;

        FISCA.Data.QueryHelper _QueryHelper = new FISCA.Data.QueryHelper();

        /// <summary>
        /// 目前僅記錄列印紙張尺寸
        /// </summary>
        public string ConfigPrint = "日常表現記錄表_新制_列印設定";

        /// <summary>
        /// XML結構之設定檔
        /// </summary>
        public string ConfigType = "日常表現記錄表_新制_假別設定";

        BackgroundWorker BGW;

        public NewForm()
        {
            InitializeComponent();
        }

        private void NewForm_Load(object sender, EventArgs e)
        {
            intSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            intSemester.Value = int.Parse(K12.Data.School.DefaultSemester);

            BGW = new BackgroundWorker();
            BGW.WorkerReportsProgress = true;
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.ProgressChanged += new ProgressChangedEventHandler(BGW_ProgressChanged);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
        }

        /// <summary>
        /// 開始列印
        /// </summary>
        private void btnSave_Click(object sender, EventArgs e)
        {
            //開始列印報表
            if (!BGW.IsBusy)
            {
                _Schoolyear = intSchoolYear.Value;
                _Semester = intSemester.Value;
                Carty_paper = checkBoxX1.Checked;
                FormStop = false;
                BGW.RunWorkerAsync();
            }
            else
            {
                MsgBox.Show("系統忙碌中,請稍後再試!!");
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //電子報表
            string time = DateTime.Now.Hour.ToString().PadLeft(2) + DateTime.Now.Minute.ToString().PadLeft(2);
            paperForClass = new SmartSchool.ePaper.ElectronicPaper(string.Format("日常表現記錄表(新制_{0})", time), _Schoolyear.ToString(), _Semester.ToString(), SmartSchool.ePaper.ViewerType.Class);

            BGW.ReportProgress(1, "取得紙張設定");
            //取得列印紙張
            int sizeIndex = GetSizeIndex();

            BGW.ReportProgress(4, "取得假別設定");
            //取得列印假別內容
            Dictionary<string, List<string>> UserType = GetUserType();


            #region 取得資料

            //取得使用者選擇班級
            BGW.ReportProgress(8, "取得所選班級");
            List<ClassRecord> allClasses = Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);

            #region 取得使用者所選擇的班級學生

            BGW.ReportProgress(12, "取得學生清單");
            string classidlist = string.Join(",", K12.Presentation.NLDPanels.Class.SelectedSource);
            StringBuilder sb = new StringBuilder();
            sb.Append("select student.id,student.ref_class_id from student ");
            sb.Append("join class on class.id=student.ref_class_id ");
            sb.Append("where student.status=1 ");
            sb.Append(string.Format("and class.id in ({0})", classidlist));

            List<string> StudentIDList = new List<string>();
            DataTable dt = _QueryHelper.Select(sb.ToString());
            foreach (DataRow row in dt.Rows)
            {
                StudentIDList.Add("" + row[0]);
            }

            BGW.ReportProgress(15, "取得學生清單");
            List<StudentRecord> allStudents = Student.SelectByIDs(StudentIDList);

            int maxStudents = 0;
            int totalStudent = allStudents.Count;

            Dictionary<string, List<StudentRecord>> classStudents = new Dictionary<string, List<StudentRecord>>();
            foreach (StudentRecord each in allStudents)
            {
                if (!classStudents.ContainsKey(each.RefClassID))
                {
                    classStudents.Add(each.RefClassID, new List<StudentRecord>());
                }

                classStudents[each.RefClassID].Add(each);
            }

            foreach (string each in classStudents.Keys)
            {
                if (classStudents[each].Count > maxStudents)
                {
                    maxStudents = classStudents[each].Count;
                }

                classStudents[each].Sort(SortStudent);
            }
            #endregion


            Dictionary<string, RewardRecord> MeritDemeritAttDic = new Dictionary<string, RewardRecord>();

            BGW.ReportProgress(20, "取得獎勵資料");
            #region 獎勵

            foreach (SHMeritRecord each in SHMerit.SelectByStudentIDs(StudentIDList))
            {
                if (each.SchoolYear != _Schoolyear || each.Semester != _Semester)
                    continue;

                RewardRecord rr = new RewardRecord();

                if (MeritDemeritAttDic.ContainsKey(each.RefStudentID))
                {
                    rr = MeritDemeritAttDic[each.RefStudentID];
                }

                rr.MeritACount += each.MeritA.HasValue ? each.MeritA.Value : 0;
                rr.MeritBCount += each.MeritB.HasValue ? each.MeritB.Value : 0;
                rr.MeritCCount += each.MeritC.HasValue ? each.MeritC.Value : 0;

                if (!MeritDemeritAttDic.ContainsKey(each.RefStudentID))
                {
                    MeritDemeritAttDic.Add(each.RefStudentID, rr);
                }
            }
            #endregion

            BGW.ReportProgress(25, "取得懲戒資料");
            #region 懲戒
            foreach (SHDemeritRecord each in SHDemerit.SelectByStudentIDs(StudentIDList))
            {
                if (each.SchoolYear != _Schoolyear || each.Semester != _Semester)
                    continue;

                if (each.Cleared == "是")
                    continue;

                RewardRecord rr = new RewardRecord();

                if (MeritDemeritAttDic.ContainsKey(each.RefStudentID))
                {
                    rr = MeritDemeritAttDic[each.RefStudentID];
                }

                rr.DemeritACount += each.DemeritA.HasValue ? each.DemeritA.Value : 0;
                rr.DemeritBCount += each.DemeritB.HasValue ? each.DemeritB.Value : 0;
                rr.DemeritCCount += each.DemeritC.HasValue ? each.DemeritC.Value : 0;

                if (!MeritDemeritAttDic.ContainsKey(each.RefStudentID))
                {
                    MeritDemeritAttDic.Add(each.RefStudentID, rr);
                }
            }
            #endregion

            BGW.ReportProgress(30, "取得節次對照");
            //取得節次對照表
            Dictionary<string, string> periodDic = new Dictionary<string, string>();
            foreach (PeriodMappingInfo var in PeriodMapping.SelectAll())
            {
                string name = var.Type;

                //取得對照表並且對照出節次->類別的清單(99/11/24 by dylan)
                if (!periodDic.ContainsKey(var.Name))
                {
                    periodDic.Add(var.Name, var.Type);
                }
            }

            BGW.ReportProgress(35, "取得缺曠資料");
            #region 缺曠
            foreach (SHAttendanceRecord each in SHAttendance.SelectByStudentIDs(StudentIDList))
            {
                if (each.SchoolYear != _Schoolyear || each.Semester != _Semester)
                    continue;

                RewardRecord rr = new RewardRecord();
                if (MeritDemeritAttDic.ContainsKey(each.RefStudentID))
                {
                    rr = MeritDemeritAttDic[each.RefStudentID];
                }

                foreach (AttendancePeriod _Period in each.PeriodDetail)
                {
                    if (periodDic.ContainsKey(_Period.Period))
                    {
                        string typename = periodDic[_Period.Period] + "_" + _Period.AbsenceType;
                        if (rr.Attendance.ContainsKey(typename))
                        {
                            rr.Attendance[typename]++;
                        }
                        else
                        {
                            rr.Attendance.Add(typename, 1);
                        }
                    }
                }

                if (!MeritDemeritAttDic.ContainsKey(each.RefStudentID))
                {
                    MeritDemeritAttDic.Add(each.RefStudentID, rr);
                }
            }
            #endregion

            //日常表現資料
            BGW.ReportProgress(40, "取得日常表現");
            Dictionary<string, SHMoralScoreRecord> SHMoralScoreDic = new Dictionary<string, SHMoralScoreRecord>();
            foreach (SHMoralScoreRecord each in SHMoralScore.Select(null, StudentIDList, _Schoolyear, _Semester))
            {
                if (!SHMoralScoreDic.ContainsKey(each.RefStudentID))
                {
                    SHMoralScoreDic.Add(each.RefStudentID, each);
                }
            }

            //文字評量對照表
            BGW.ReportProgress(45, "取得文字評量");
            List<string> TextScoreList = new List<string>();
            SmartSchool.Customization.Data.SystemInformation.getField("文字評量對照表");
            System.Xml.XmlElement ElmTextScoreList = (System.Xml.XmlElement)SmartSchool.Customization.Data.SystemInformation.Fields["文字評量對照表"];
            foreach (System.Xml.XmlNode Node in ElmTextScoreList.SelectNodes("Content/Morality"))
                TextScoreList.Add(Node.Attributes["Face"].InnerText);

            #endregion

            #region 產生表格

            BGW.ReportProgress(50, "產生報表樣式");

            Workbook template = new Workbook();
            Workbook prototype = new Workbook();

            //列印尺寸
            if (sizeIndex == 0)
                template.Open(new MemoryStream(Properties.Resources.德行表現總表新制A3), FileFormatType.Excel2003);
            else if (sizeIndex == 1)
                template.Open(new MemoryStream(Properties.Resources.德行表現總表新制A4), FileFormatType.Excel2003);
            else if (sizeIndex == 2)
                template.Open(new MemoryStream(Properties.Resources.德行表現總表新制B4), FileFormatType.Excel2003);

            prototype.Copy(template);

            Worksheet templateSheet = template.Worksheets[0];
            Worksheet prototypeSheet = prototype.Worksheets[0];

            Range tempAbsence = templateSheet.Cells.CreateRange(9, 1, true);
            Range tempScoreText = templateSheet.Cells.CreateRange(10, 1, true);
            Range tempAfterOtherDiff = templateSheet.Cells.CreateRange(11, 1, true);
            Range oder = templateSheet.Cells.CreateRange(12, 1, true);

            Dictionary<string, int> columnIndexTable = new Dictionary<string, int>();

            Dictionary<string, List<string>> periodAbsence = new Dictionary<string, List<string>>();

            //紀錄獎懲的 Column Index
            columnIndexTable.Add("大功", 3);
            columnIndexTable.Add("小功", 4);
            columnIndexTable.Add("嘉獎", 5);
            columnIndexTable.Add("大過", 6);
            columnIndexTable.Add("小過", 7);
            columnIndexTable.Add("警告", 8);

            //缺曠
            int ptColIndex = 9;
            foreach (string var in UserType.Keys)
            {
                foreach (string absence in UserType[var])
                {
                    if (!periodAbsence.ContainsKey(var))
                        periodAbsence.Add(var, new List<string>());
                    if (!periodAbsence[var].Contains(absence))
                        periodAbsence[var].Add(absence);

                    prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(tempAbsence);
                    ptColIndex += 1;
                }
            }

            ptColIndex = 9;
            foreach (string period in periodAbsence.Keys)
            {
                prototypeSheet.Cells.CreateRange(2, ptColIndex, 1, periodAbsence[period].Count).Merge();
                prototypeSheet.Cells[2, ptColIndex].PutValue(period);

                foreach (string absence in periodAbsence[period])
                {
                    prototypeSheet.Cells[3, ptColIndex].PutValue(absence);
                    columnIndexTable.Add(period + "_" + absence, ptColIndex);
                    ptColIndex++;
                }
            }

            if (ptColIndex > 9)
            {
                prototypeSheet.Cells.CreateRange(1, 9, 1, ptColIndex - 9).Merge();
                prototypeSheet.Cells[1, 9].PutValue("缺曠");
            }

            //用來調整Column寬度的定位
            int ColumnMax = ptColIndex;

            //文字評量
            foreach (string textscore in TextScoreList)
            {
                columnIndexTable.Add(textscore, ptColIndex);
                prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(tempScoreText);
                prototypeSheet.Cells[4, ptColIndex].PutValue(textscore);
                ptColIndex++;
            }

            prototypeSheet.Cells[1, ptColIndex - TextScoreList.Count].PutValue("學生綜合表現");

            if ((ptColIndex - TextScoreList.Count > 0) && (TextScoreList.Count > 0))
            {
                prototypeSheet.Cells.CreateRange(1, ptColIndex - TextScoreList.Count, 1, TextScoreList.Count).Merge();
                prototypeSheet.Cells.CreateRange(2, ptColIndex - TextScoreList.Count, 1, TextScoreList.Count).Merge();
                prototypeSheet.Cells.CreateRange(3, ptColIndex - TextScoreList.Count, 1, TextScoreList.Count).Merge();
            }

            prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(tempAfterOtherDiff);
            columnIndexTable.Add("評語", ptColIndex++);
            prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(oder);
            columnIndexTable.Add("是否留察", ptColIndex++);

            //填入製表日期
            prototypeSheet.Cells[0, 0].PutValue("製表日期：" + DateTime.Today.ToShortDateString());

            //填入標題
            prototypeSheet.Cells.CreateRange(0, 3, 1, ptColIndex - 3).Merge();
            prototypeSheet.Cells[0, 3].PutValue(K12.Data.School.ChineseName + " " + _Schoolyear + " 學年度 " + ((_Semester == 1) ? "上" : "下") + " 學期 日常表現記錄表（新制）");

            Range ptEachRow = prototypeSheet.Cells.CreateRange(5, 1, false);

            for (int i = 5; i < maxStudents + 5; i++)
            {
                prototypeSheet.Cells.CreateRange(i, 1, false).Copy(ptEachRow);
            }

            //加上底線
            prototypeSheet.Cells.CreateRange(maxStudents + 5, 0, 1, ptColIndex).SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, System.Drawing.Color.Black);

            for (int i = 12; i >= ptColIndex; i--)
                prototypeSheet.Cells.DeleteColumn(i);

            Range pt = prototypeSheet.Cells.CreateRange(0, maxStudents + 5, false);

            #endregion

            #region 填入表格

            BGW.ReportProgress(53, "填入報表資料");


            Workbook wb = new Workbook();
            wb.Copy(prototype);
            Worksheet ws = wb.Worksheets[0];

            int index = 0;
            int dataIndex = 0;
            int classTotalRow = maxStudents + 5;

            BGW.ReportProgress(57, "填入老師姓名");
            #region 取得全校班級,與老師姓名/暱稱(2012/5/24)
            Dictionary<string, string> TeacherDic = new Dictionary<string, string>();
            string st = "SELECT class.id,teacher.teacher_name,teacher.nickname FROM class JOIN teacher ON class.ref_teacher_id = teacher.id";
            dt = _QueryHelper.Select(st);
            foreach (DataRow row in dt.Rows)
            {
                string classID = "" + row[0];
                string teacherName = "" + row[1];
                string teacherNickname = "" + row[2];

                if (!TeacherDic.ContainsKey(classID))
                {
                    if (string.IsNullOrEmpty(teacherNickname))
                        TeacherDic.Add(classID, teacherName);
                    else
                        TeacherDic.Add(classID, teacherName + "(" + teacherNickname + ")");
                }
            }

            #endregion

            BGW.ReportProgress(61, "學生留察資料");
            #region 取得全校本學年度留查之記錄(2012/5/24)
            List<string> MeritFlagIs2 = new List<string>();
            st = string.Format("SELECT ref_student_id from discipline where merit_flag=2 and school_year={0} and semester={1}", _Schoolyear, _Semester);
            dt = _QueryHelper.Select(st);
            foreach (DataRow row in dt.Rows)
            {
                if (!MeritFlagIs2.Contains("" + row[0]))
                {
                    MeritFlagIs2.Add("" + row[0]);
                }
            }

            #endregion

            BGW.ReportProgress(70, "開始列印資料");

            int PeogressNow1 = totalStudent / 30;
            int PeogressNow2 = 0;
            int PeogressNow3 = 70;

            foreach (ClassRecord aClass in allClasses)
            {
                //電子報表用
                #region 電子報表用

                Workbook Paper_wb = new Workbook();
                Paper_wb.Copy(prototype);
                Worksheet Paper_ws = Paper_wb.Worksheets[0];

                int Paper_index = 0;
                int Paper_dataIndex = 0;
                int Paper_classTotalRow = maxStudents + 5;

                int Paper_PeogressNow1 = totalStudent / 30;


                #endregion

                if (!classStudents.ContainsKey(aClass.ID))
                    continue;

                string TeacherName = "";
                if (TeacherDic.ContainsKey(aClass.ID))
                {
                    TeacherName = TeacherDic[aClass.ID];
                }

                //複製完成後的樣板
                ws.Cells.CreateRange(index, classTotalRow, false).Copy(pt);
                Paper_ws.Cells.CreateRange(Paper_index, Paper_classTotalRow, false).Copy(pt);

                //填入班級名稱
                ws.Cells[index + 1, 0].PutValue("班級：" + aClass.Name);
                Paper_ws.Cells[Paper_index + 1, 0].PutValue("班級：" + aClass.Name);

                //填入老師名稱
                ws.Cells[index + 3, 0].PutValue("教師：" + TeacherName);
                Paper_ws.Cells[Paper_index + 3, 0].PutValue("教師：" + TeacherName);

                dataIndex = index + 5;
                Paper_dataIndex = Paper_index + 5;

                foreach (StudentRecord aStudent in classStudents[aClass.ID])
                {
                    PeogressNow2++;

                    if (PeogressNow2 > PeogressNow1 && PeogressNow3 < 101)
                    {
                        PeogressNow3++;
                        PeogressNow2 = 0;
                        BGW.ReportProgress(PeogressNow3, "開始列印資料");
                    }

                    ws.Cells[dataIndex, 0].PutValue(aStudent.SeatNo);
                    ws.Cells[dataIndex, 1].PutValue(aStudent.Name);
                    ws.Cells[dataIndex, 2].PutValue(aStudent.StudentNumber);

                    Paper_ws.Cells[Paper_dataIndex, 0].PutValue(aStudent.SeatNo);
                    Paper_ws.Cells[Paper_dataIndex, 1].PutValue(aStudent.Name);
                    Paper_ws.Cells[Paper_dataIndex, 2].PutValue(aStudent.StudentNumber);

                    if (MeritDemeritAttDic.ContainsKey(aStudent.ID))
                    {
                        RewardRecord rr = MeritDemeritAttDic[aStudent.ID];

                        ws.Cells[dataIndex, columnIndexTable["大功"]].PutValue(rr.MeritACount != 0 ? rr.MeritACount.ToString() : "");
                        ws.Cells[dataIndex, columnIndexTable["小功"]].PutValue(rr.MeritBCount != 0 ? rr.MeritBCount.ToString() : "");
                        ws.Cells[dataIndex, columnIndexTable["嘉獎"]].PutValue(rr.MeritCCount != 0 ? rr.MeritCCount.ToString() : "");
                        ws.Cells[dataIndex, columnIndexTable["大過"]].PutValue(rr.DemeritACount != 0 ? rr.DemeritACount.ToString() : "");
                        ws.Cells[dataIndex, columnIndexTable["小過"]].PutValue(rr.DemeritBCount != 0 ? rr.DemeritBCount.ToString() : "");
                        ws.Cells[dataIndex, columnIndexTable["警告"]].PutValue(rr.DemeritCCount != 0 ? rr.DemeritCCount.ToString() : "");

                        Paper_ws.Cells[Paper_dataIndex, columnIndexTable["大功"]].PutValue(rr.MeritACount != 0 ? rr.MeritACount.ToString() : "");
                        Paper_ws.Cells[Paper_dataIndex, columnIndexTable["小功"]].PutValue(rr.MeritBCount != 0 ? rr.MeritBCount.ToString() : "");
                        Paper_ws.Cells[Paper_dataIndex, columnIndexTable["嘉獎"]].PutValue(rr.MeritCCount != 0 ? rr.MeritCCount.ToString() : "");
                        Paper_ws.Cells[Paper_dataIndex, columnIndexTable["大過"]].PutValue(rr.DemeritACount != 0 ? rr.DemeritACount.ToString() : "");
                        Paper_ws.Cells[Paper_dataIndex, columnIndexTable["小過"]].PutValue(rr.DemeritBCount != 0 ? rr.DemeritBCount.ToString() : "");
                        Paper_ws.Cells[Paper_dataIndex, columnIndexTable["警告"]].PutValue(rr.DemeritCCount != 0 ? rr.DemeritCCount.ToString() : "");

                        foreach (string each in rr.Attendance.Keys)
                        {
                            if (columnIndexTable.ContainsKey(each))
                            {
                                ws.Cells[dataIndex, columnIndexTable[each]].PutValue(rr.Attendance[each]);
                                Paper_ws.Cells[Paper_dataIndex, columnIndexTable[each]].PutValue(rr.Attendance[each]);
                            }
                        }
                    }

                    //文字評量部份
                    SHMoralScoreRecord demonScore;
                    if (SHMoralScoreDic.ContainsKey(aStudent.ID))
                    {
                        demonScore = SHMoralScoreDic[aStudent.ID];

                        //文字評量
                        XmlElement xml = demonScore.TextScore;
                        foreach (XmlElement each in xml.SelectNodes("Morality"))
                        {
                            string strFace = each.GetAttribute("Face");
                            if (columnIndexTable.ContainsKey(strFace))
                            {
                                int colIndex = columnIndexTable[strFace];
                                ws.Cells[dataIndex, colIndex].PutValue(each.InnerText);
                                Paper_ws.Cells[Paper_dataIndex, colIndex].PutValue(each.InnerText);
                            }
                        }

                        //導師評語
                        ws.Cells[dataIndex, columnIndexTable["評語"]].PutValue(demonScore.Comment);
                        Paper_ws.Cells[Paper_dataIndex, columnIndexTable["評語"]].PutValue(demonScore.Comment);
                    }

                    //留察
                    if (MeritFlagIs2.Contains(aStudent.ID))
                    {
                        ws.Cells[dataIndex, columnIndexTable["是否留察"]].PutValue("是");
                        Paper_ws.Cells[Paper_dataIndex, columnIndexTable["是否留察"]].PutValue("是");
                    }

                    foreach (int each in columnIndexTable.Values)
                    {
                        if (each >= ColumnMax)
                        {
                            Paper_ws.AutoFitColumn(each);
                        }
                    }

                    dataIndex++;
                    Paper_dataIndex++;
                }

                //電子報表
                MemoryStream stream = Paper_wb.SaveToStream();
                paperForClass.Append(new PaperItem(PaperFormat.Office2003Xls, stream, aClass.ID));

                index += classTotalRow + 2;
                ws.HPageBreaks.Add(index, ptColIndex);
            }

            foreach (int each in columnIndexTable.Values)
            {
                if (each >= ColumnMax)
                {
                    ws.AutoFitColumn(each);
                }
            }
            #endregion
            if (Carty_paper)
            {
                BGW.ReportProgress(90, "上傳電子報表");
                SmartSchool.ePaper.DispatcherProvider.Dispatch(paperForClass);
            }

            BGW.ReportProgress(100, "資料列印完成");
            e.Result = wb;
        }

        void BGW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState, e.ProgressPercentage);
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("日常表現記錄表（新制）列印完成");
                Completed.Save("日常表現記錄表（新制）", (Workbook)e.Result);
            }
            else
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("日常表現記錄表（新制）,報表列印失敗...\n" + e.Error.Message);
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
            }
            FormStop = true;
        }

        //紙張設定
        private int GetSizeIndex()
        {
            Campus.Configuration.ConfigData cd = Campus.Configuration.Config.User[ConfigPrint];
            string config = cd["紙張設定"];
            int x = 0;
            int.TryParse(config, out x);
            return x; //如果是數值就回傳,如果不是回傳預設
        }

        private int SortStudent(StudentRecord sr1, StudentRecord sr2)
        {
            int srr1 = sr1.SeatNo.HasValue ? sr1.SeatNo.Value : 0;
            int srr2 = sr2.SeatNo.HasValue ? sr2.SeatNo.Value : 0;
            return srr1.CompareTo(srr2);
        }

        //節次設定
        private Dictionary<string, List<string>> GetUserType()
        {
            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            Campus.Configuration.ConfigData cd = Campus.Configuration.Config.User[ConfigType];
            string config = cd["節次設定"];

            if (!string.IsNullOrEmpty(config))
            {
                XmlElement print = DSXmlHelper.LoadXml(config);

                foreach (XmlElement type in print.SelectNodes("Type"))
                {
                    string typeName = type.GetAttribute("Text");

                    if (!dic.ContainsKey(typeName))
                        dic.Add(typeName, new List<string>());

                    foreach (XmlElement absence in type.SelectNodes("Absence"))
                    {
                        string absenceName = absence.GetAttribute("Text");

                        if (!dic[typeName].Contains(absenceName))
                            dic[typeName].Add(absenceName);
                    }
                }
            }


            return dic;
        }

        /// <summary>
        /// 列印紙張設定
        /// </summary>
        private void linkPrint_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SelectPrintSizeForm config = new SelectPrintSizeForm(ConfigPrint);
            config.ShowDialog();
        }

        /// <summary>
        /// 列印節次設定
        /// </summary>
        private void linkType_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SelectTypeForm typeForm = new SelectTypeForm(ConfigType);
            typeForm.ShowDialog();
        }

        bool FormStop
        {
            set
            {
                intSchoolYear.Enabled = value;
                intSemester.Enabled = value;
                linkPrint.Enabled = value;
                linkType.Enabled = value;
                btnSave.Enabled = value;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
