using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.DSAUtil;
using System.Xml;
using System.IO;
using K12.Data;
using SHSchool.Data;
using SmartSchool.ePaper;
using Aspose.Words;
using System.Diagnostics;
using SmartSchool.Feature.Student;

namespace 德行成績試算表
{
    public partial class NewDocxForm : BaseForm
    {
        ////主文件
        private Document _doc;

        ////單頁範本
        //private Document _template;
        ////
        //private byte[] _buffer = null;

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

        public NewDocxForm()
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

            _doc = new Document();
            _doc.Sections.Clear(); //清空此Document

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
            //排序(因為上面沒有照班級排序)
            int displayOrder;
            allClasses = allClasses.OrderBy(i => i.GradeYear).ThenBy(i => displayOrder = Int32.TryParse(i.DisplayOrder, out displayOrder) ? displayOrder : 0).ThenBy(i => i.Name).ToList();



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
            List<StudentRecord> allStudents = K12.Data.Student.SelectByIDs(StudentIDList);

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
            Dictionary<string, RewardRecord> TotalMeritDemeritDic = new Dictionary<string, RewardRecord>();

            BGW.ReportProgress(20, "取得獎勵資料");
            #region 獎勵

            foreach (SHMeritRecord each in SHMerit.SelectByStudentIDs(StudentIDList))
            {
                // 2018/1/8 羿均 新增 累計獎勵紀錄資料
                RewardRecord totalMerit = new RewardRecord();
                if (TotalMeritDemeritDic.ContainsKey(each.RefStudentID))
                {
                    totalMerit = TotalMeritDemeritDic[each.RefStudentID];
                }

                totalMerit.MeritACount += each.MeritA.HasValue ? each.MeritA.Value : 0;
                totalMerit.MeritBCount += each.MeritB.HasValue ? each.MeritB.Value : 0;
                totalMerit.MeritCCount += each.MeritC.HasValue ? each.MeritC.Value : 0;

                if (!TotalMeritDemeritDic.ContainsKey(each.RefStudentID))
                {
                    TotalMeritDemeritDic.Add(each.RefStudentID, totalMerit);
                }

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
                // 2018/1/8 羿均 新增 累計懲戒紀錄資料
                RewardRecord totalDemerit = new RewardRecord();

                if (each.Cleared == "是")
                    continue;

                if (TotalMeritDemeritDic.ContainsKey(each.RefStudentID))
                {
                    totalDemerit = TotalMeritDemeritDic[each.RefStudentID];
                }

                totalDemerit.DemeritACount += each.DemeritA.HasValue ? each.DemeritA.Value : 0;
                totalDemerit.DemeritBCount += each.DemeritB.HasValue ? each.DemeritB.Value : 0;
                totalDemerit.DemeritCCount += each.DemeritC.HasValue ? each.DemeritC.Value : 0;

                if (!TotalMeritDemeritDic.ContainsKey(each.RefStudentID))
                {
                    TotalMeritDemeritDic.Add(each.RefStudentID, totalDemerit);
                }

                if (each.SchoolYear != _Schoolyear || each.Semester != _Semester)
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
                        string typePeriod = periodDic[_Period.Period];
                        if (!rr.typeAttendance.ContainsKey(typePeriod))
                        {
                            rr.typeAttendance.Add(typePeriod, new Dictionary<string, int>());
                        }
                        if (!rr.typeAttendance[typePeriod].ContainsKey(_Period.AbsenceType))
                        {
                            rr.typeAttendance[typePeriod].Add(_Period.AbsenceType, 0);
                        }
                        rr.typeAttendance[typePeriod][_Period.AbsenceType]++;

                        string typename = periodDic[_Period.Period] + _Period.AbsenceType;
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
            {
                TextScoreList.Add(Node.Attributes["Face"].InnerText);
            }

            #endregion

            #region 產生表格

            BGW.ReportProgress(50, "產生報表樣式");

            //單頁範本
            Document _template = new Document(new MemoryStream(Properties.Resources.高中_日常生活表現_班級_a3));

            Dictionary<string, List<string>> periodAbsence = new Dictionary<string, List<string>>();


            //New
            List<string> nameList = new List<string>();
            List<string> valueList = new List<string>();
            nameList.Add("學校名稱");
            valueList.Add(School.ChineseName);
            nameList.Add("列印日期");
            valueList.Add(DateTime.Today.ToString("yyyy/MM/dd"));

            nameList.Add("學年度");
            valueList.Add("" + _Schoolyear);

            nameList.Add("學期");
            valueList.Add(((_Semester == 1) ? "上" : "下") + " 學期");


            int nameIndex = 1;
            int varIndex = 1;
            foreach (string var in UserType.Keys)
            {
                //New
                nameList.Add("節次類型" + varIndex);
                valueList.Add(var);

                foreach (string absence in UserType[var])
                {
                    if (!periodAbsence.ContainsKey(var))
                        periodAbsence.Add(var, new List<string>());
                    if (!periodAbsence[var].Contains(absence))
                        periodAbsence[var].Add(absence);

                    //New
                    nameList.Add("節次類型" + varIndex + "缺曠" + nameIndex);
                    valueList.Add(var + absence);
                    nameIndex++;
                }
                varIndex++;
            }

            //foreach (string period in periodAbsence.Keys)
            //{
            //    prototypeSheet.Cells.CreateRange(2, ptColIndex, 2, periodAbsence[period].Count).Merge();
            //    prototypeSheet.Cells[2, ptColIndex].PutValue(period);
            //
            //    foreach (string absence in periodAbsence[period])
            //    {
            //        prototypeSheet.Cells[4, ptColIndex].PutValue(absence);
            //        columnIndexTable.Add(period + absence, ptColIndex);
            //        ptColIndex++;
            //    }
            //}

            //文字評量
            int textscoreIndex = 1;
            foreach (string textscore in TextScoreList)
            {
                //New
                nameList.Add("評語標題" + textscoreIndex);
                valueList.Add(textscore);
                textscoreIndex++;
            }

            #endregion

            #region 填入表格

            BGW.ReportProgress(53, "填入報表資料");


            int index = 0;
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



            foreach (ClassRecord aClass in allClasses)
            {
                //每一張每一份
                Document PageOne = (Document)_template.Clone(true);

                if (!classStudents.ContainsKey(aClass.ID))
                    continue;

                string TeacherName = "";
                if (TeacherDic.ContainsKey(aClass.ID))
                {
                    TeacherName = TeacherDic[aClass.ID];
                }

                nameList.Add("班級");
                valueList.Add(aClass.Name);

                nameList.Add("教師");
                valueList.Add(TeacherName);

                int aStudentIndex = 1;
                foreach (StudentRecord aStudent in classStudents[aClass.ID])
                {

                    nameList.Add("座號" + aStudentIndex);
                    valueList.Add("" + aStudent.SeatNo);
                    nameList.Add("姓名" + aStudentIndex);
                    valueList.Add(aStudent.Name);
                    nameList.Add("學號" + aStudentIndex);
                    valueList.Add(aStudent.StudentNumber);

                    // 2018/01/17 羿均
                    // MeritDemeritAttDic 為當學年度學期 缺曠獎懲資料
                    // TotalMeritDemeritDic 為獎懲累計資料
                    if (MeritDemeritAttDic.ContainsKey(aStudent.ID) || TotalMeritDemeritDic.ContainsKey(aStudent.ID))
                    {
                        RewardRecord rr = new RewardRecord();
                        if (MeritDemeritAttDic.ContainsKey(aStudent.ID))
                        {
                            rr = MeritDemeritAttDic[aStudent.ID];
                        }
                        if (TotalMeritDemeritDic.ContainsKey(aStudent.ID))
                        {
                            RewardRecord totalRecord = TotalMeritDemeritDic[aStudent.ID];

                            nameList.Add("大功學期" + aStudentIndex);
                            valueList.Add("" + rr.MeritACount);
                            nameList.Add("大功累積" + aStudentIndex);
                            valueList.Add("" + totalRecord.MeritACount);

                            nameList.Add("小功學期" + aStudentIndex);
                            valueList.Add("" + rr.MeritBCount);
                            nameList.Add("小功累積" + aStudentIndex);
                            valueList.Add("" + totalRecord.MeritBCount);

                            nameList.Add("嘉獎學期" + aStudentIndex);
                            valueList.Add("" + rr.MeritCCount);
                            nameList.Add("嘉獎累積" + aStudentIndex);
                            valueList.Add("" + totalRecord.MeritCCount);

                            nameList.Add("大過學期" + aStudentIndex);
                            valueList.Add("" + rr.DemeritACount);
                            nameList.Add("大過累積" + aStudentIndex);
                            valueList.Add("" + totalRecord.DemeritACount);

                            nameList.Add("小過學期" + aStudentIndex);
                            valueList.Add("" + rr.DemeritBCount);
                            nameList.Add("小過累積" + aStudentIndex);
                            valueList.Add("" + totalRecord.DemeritBCount);

                            nameList.Add("警告學期" + aStudentIndex);
                            valueList.Add("" + rr.DemeritCCount);
                            nameList.Add("警告累積" + aStudentIndex);
                            valueList.Add("" + totalRecord.DemeritCCount);

                        }
                        else
                        {
                            //有缺曠沒獎懲，要印獎懲0/0
                            //https://3.basecamp.com/4399967/buckets/15765350/todos/4520462594

                            nameList.Add("大功學期" + aStudentIndex);
                            valueList.Add("0");
                            nameList.Add("大功累積" + aStudentIndex);
                            valueList.Add("0");

                            nameList.Add("小功學期" + aStudentIndex);
                            valueList.Add("0");
                            nameList.Add("小功累積" + aStudentIndex);
                            valueList.Add("0");

                            nameList.Add("嘉獎學期" + aStudentIndex);
                            valueList.Add("0");
                            nameList.Add("嘉獎累積" + aStudentIndex);
                            valueList.Add("0");

                            nameList.Add("大過學期" + aStudentIndex);
                            valueList.Add("0");
                            nameList.Add("大過累積" + aStudentIndex);
                            valueList.Add("0");

                            nameList.Add("小過學期" + aStudentIndex);
                            valueList.Add("0");
                            nameList.Add("小過累積" + aStudentIndex);
                            valueList.Add("0");

                            nameList.Add("警告學期" + aStudentIndex);
                            valueList.Add("0");
                            nameList.Add("警告累積" + aStudentIndex);
                            valueList.Add("0");
                        }

                        //缺曠
                        int typeIndex = 1;
                        foreach (string type in rr.typeAttendance.Keys)
                        {
                            int attendIndex = 1;
                            foreach (string attend in rr.typeAttendance[type].Keys)
                            {
                                nameList.Add("節次類型" + typeIndex + "缺曠" + attendIndex + "學生" + aStudentIndex);
                                valueList.Add("" + rr.typeAttendance[type][attend]);
                                attendIndex++;
                            }
                            typeIndex++;
                        }
                    }
                    else
                    {
                        //沒缺曠沒獎懲，要印獎懲0/0
                        // https://3.basecamp.com/4399967/buckets/15765350/todos/4520462594
                        nameList.Add("大功學期" + aStudentIndex);
                        valueList.Add("0");
                        nameList.Add("大功累積" + aStudentIndex);
                        valueList.Add("0");

                        nameList.Add("小功學期" + aStudentIndex);
                        valueList.Add("0");
                        nameList.Add("小功累積" + aStudentIndex);
                        valueList.Add("0");

                        nameList.Add("嘉獎學期" + aStudentIndex);
                        valueList.Add("0");
                        nameList.Add("嘉獎累積" + aStudentIndex);
                        valueList.Add("0");

                        nameList.Add("大過學期" + aStudentIndex);
                        valueList.Add("0");
                        nameList.Add("大過累積" + aStudentIndex);
                        valueList.Add("0");

                        nameList.Add("小過學期" + aStudentIndex);
                        valueList.Add("0");
                        nameList.Add("小過累積" + aStudentIndex);
                        valueList.Add("0");

                        nameList.Add("警告學期" + aStudentIndex);
                        valueList.Add("0");
                        nameList.Add("警告累積" + aStudentIndex);
                        valueList.Add("0");
                    }

                    //文字評量部份
                    SHMoralScoreRecord demonScore;

                    if (SHMoralScoreDic.ContainsKey(aStudent.ID))
                    {
                        demonScore = SHMoralScoreDic[aStudent.ID];

                        //文字評量
                        XmlElement xml = demonScore.TextScore;
                        int demonScoreIndex = 1;
                        foreach (XmlElement each in xml.SelectNodes("Morality"))
                        {
                            string strFace = each.GetAttribute("Face");

                            if (TextScoreList.Contains(strFace))
                            {
                                nameList.Add("評語" + demonScoreIndex + "學生" + aStudentIndex);
                                valueList.Add(each.InnerText);
                                demonScoreIndex++;
                            }
                        }

                        //導師評語
                        nameList.Add("導師評語" + aStudentIndex);
                        valueList.Add(demonScore.Comment);
                    }

                    //留察
                    if (MeritFlagIs2.Contains(aStudent.ID))
                    {
                        nameList.Add("是否留察" + aStudentIndex);
                        valueList.Add("是");
                    }

                    //下一位學生
                    aStudentIndex++;
                }

                PageOne.MailMerge.Execute(nameList.ToArray(), valueList.ToArray());
                PageOne.MailMerge.DeleteFields();

                //電子報表
                MemoryStream stream = new MemoryStream();
                PageOne.Save(stream, SaveFormat.Docx);
                paperForClass.Append(new PaperItem(PaperFormat.Office2003Doc, stream, aClass.ID));

                _doc.Sections.Add(_doc.ImportNode(PageOne.FirstSection, true));
            }

            #endregion
            if (Carty_paper)
            {
                BGW.ReportProgress(90, "上傳電子報表");
                SmartSchool.ePaper.DispatcherProvider.Dispatch(paperForClass);
            }

            BGW.ReportProgress(100, "資料列印完成");
            e.Result = _doc;
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
                Completed.SaveDoc("日常表現記錄表（新制）", (Document)e.Result);
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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = string.Format("日常生活表現_班級_功能變數總表_{0}.docx", DateTime.Now.ToString("HHmmss"));
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.日常生活表現總表_功能變數總表, 0, Properties.Resources.日常生活表現總表_功能變數總表.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }
    }
}
