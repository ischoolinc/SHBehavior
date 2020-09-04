using Aspose.Cells;
using FISCA.Presentation.Controls;
using K12.Data;
using K12.Data.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
//using Campus.ePaperCloud;
using FISCA.Authentication;

namespace 日常生活表現預警表
{
    public partial class PerformanceAlertForm : BaseForm
    {

        string ConfigName1 = "日常生活表現預警表_缺曠別設定_採計節次";

        string ConfigName2 = "日常生活表現預警表_畫面設定";
        string ConfigName2_1 = "採計節次";
        string ConfigName2_2 = "懲戒採計";

        string ConfigName2_相抵 = "是否功過相抵";
        string ConfigName2_換算 = "是否功過換算";
        string ConfigName2_電子報 = "是否發送教師電子報表";
        string ConfigName2_預設推播 = "預設推播訊息";

        /// <summary>
        /// 是否功過換算
        /// </summary>
        bool _IsAtoB換算 = false;

        /// <summary>
        /// 是否功過相抵
        /// </summary>
        bool _IsAtoB相抵 = false;
        bool _IsePaper電子報 = false;

        string _MessageName = "您的電子報表已收到最新內容<br>「{0}學年度 第{1}學期 日常生活表現預警表」";

        string ConfigName3 = "日常生活表現預警表_缺曠別設定_顯示節次";

        int SheetIndex = 0;

        BackgroundWorker bgw = new BackgroundWorker();
        BackgroundWorker bgw_load = new BackgroundWorker();
        BackgroundWorker bgw_ConfigSave = new BackgroundWorker();

        Dictionary<string, List<string>> AbsenceConfig1;
        Dictionary<string, List<string>> AbsenceConfig2;
        private Dictionary<string, int> ColumnInTitleIndex;

        Range RangeDetil;
        Range RangeRow;

        private Dictionary<string, string> periodDic { get; set; }

        private Workbook wb;

        private List<string> SendMessageTeacher { get; set; }

        //建立一份電子報表版本
        private Workbook wb_clsss;

        private Style sy { get; set; }

        /// <summary>
        /// 學年度
        /// </summary>
        string _SchoolYear = "107";

        /// <summary>
        /// 學期
        /// </summary>
        string _Semester = "1";

        /// <summary>
        /// 採計節次
        /// </summary>
        int _MiningSection = 0;

        /// <summary>
        /// 懲戒採計
        /// </summary>
        int _DisciplinaryMeasures = 0;

        MeritDemeritReduceRecord reduceRecord;

        //ePaperCloud ePaperCloud = new ePaperCloud();

        public PerformanceAlertForm()
        {
            InitializeComponent();
        }

        private void PerformanceAlertForm_Load(object sender, EventArgs e)
        {
            SetloadForm = false;

            bgw.RunWorkerCompleted += Bgw_RunWorkerCompleted;
            bgw.DoWork += Bgw_DoWork;
            bgw.ProgressChanged += bgw_load_ProgressChanged;
            bgw.WorkerReportsProgress = true;

            bgw_load.RunWorkerCompleted += Bgw_load_RunWorkerCompleted;
            bgw_load.DoWork += Bgw_load_DoWork;

            bgw_load.RunWorkerAsync();

            if (DSAServices.AccountType == AccountType.Greening)
            {
                //Greening帳號才能發送推播
                linkViewContent.Visible = true;
            }
        }

        private void bgw_load_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("列印「日常生活表現預警表」" + e.UserState, e.ProgressPercentage);
        }

        private void Bgw_load_DoWork(object sender, DoWorkEventArgs e)
        {
            //畫面開啟
            //1.取得目前介面的主設定值 ConfigName2
            LoadPreference();

            //2.取得顯示節次
            AbsenceConfig1 = LoadAbsence(ConfigName1);

            //取得採記節次
            AbsenceConfig2 = LoadAbsence(ConfigName3);

            //取得節次對照表設定
            periodDic = new Dictionary<string, string>();
            foreach (PeriodMappingInfo each in K12.Data.PeriodMapping.SelectAll())
            {
                if (!periodDic.ContainsKey(each.Name))
                    periodDic.Add(each.Name, each.Type);
            }

            //預設學年度學期
            _SchoolYear = K12.Data.School.DefaultSchoolYear;
            _Semester = K12.Data.School.DefaultSemester;
        }


        /// <summary>
        /// 取得缺曠列印設定
        /// </summary>
        private Dictionary<string, List<string>> LoadAbsence(string ConfigName)
        {
            Dictionary<string, List<string>> AbsenceDic = new Dictionary<string, List<string>>();

            ConfigData cd = K12.Data.School.Configuration[ConfigName];
            XmlElement preferenceData = cd.GetXml("XmlData", null);

            if (preferenceData != null)
            {
                foreach (XmlElement type in preferenceData.SelectNodes("Type"))
                {
                    string prefix = type.GetAttribute("Text");
                    if (!AbsenceDic.ContainsKey(prefix))
                        AbsenceDic.Add(prefix, new List<string>());

                    foreach (XmlElement absence in type.SelectNodes("Absence"))
                    {
                        if (!AbsenceDic[prefix].Contains(absence.GetAttribute("Text")))
                            AbsenceDic[prefix].Add(absence.GetAttribute("Text"));
                    }
                }
            }

            return AbsenceDic;

        }

        private void LoadPreference()
        {
            ConfigData cd = K12.Data.School.Configuration[ConfigName2];
            if (cd[ConfigName2_1] != "")
                int.TryParse("" + cd[ConfigName2_1], out _MiningSection);

            if (cd[ConfigName2_2] != "")
                int.TryParse("" + cd[ConfigName2_2], out _DisciplinaryMeasures);

            if (cd[ConfigName2_相抵] != "")
                bool.TryParse("" + cd[ConfigName2_相抵], out _IsAtoB相抵);

            if (cd[ConfigName2_換算] != "")
                bool.TryParse("" + cd[ConfigName2_換算], out _IsAtoB換算);

            if (cd[ConfigName2_電子報] != "")
                bool.TryParse("" + cd[ConfigName2_電子報], out _IsePaper電子報);

            if (cd[ConfigName2_預設推播] != "")
                _MessageName = "" + cd[ConfigName2_預設推播];
        }

        private void Bgw_load_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SetloadForm = true;
            if (e.Error == null)
            {
                if (!e.Cancelled)
                {
                    this.cbAtoBSum.CheckedChanged -= new System.EventHandler(this.cbAtoB換算_CheckedChanged);
                    this.cbIsAndBChange.CheckedChanged -= new System.EventHandler(this.cbIsAtoB相抵_CheckedChanged);
                    this.intMiningSection.ValueChanged -= new System.EventHandler(this.intMiningSection_ValueChanged);
                    this.intDisciplinaryMeasures.ValueChanged -= new System.EventHandler(this.intDisciplinaryMeasures_ValueChanged);
                    this.cbIsePaper.CheckedChanged -= new System.EventHandler(this.cbIsePaper_CheckedChanged);

                    integerInput1.Text = _SchoolYear; //學年度
                    integerInput2.Text = _Semester; //學期

                    intMiningSection.Value = _MiningSection;
                    intDisciplinaryMeasures.Value = _DisciplinaryMeasures;
                    cbAtoBSum.Checked = _IsAtoB換算;
                    cbIsAndBChange.Checked = _IsAtoB相抵;
                    cbIsePaper.Checked = _IsePaper電子報;

                    cbAtoBSum.Enabled = !cbIsAndBChange.Checked;
                    cbAtoBSum.Checked = cbIsAndBChange.Checked ? true : false;

                    //調整文字顏色
                    ChangeLabel();

                    this.cbAtoBSum.CheckedChanged += new System.EventHandler(this.cbAtoB換算_CheckedChanged);
                    this.cbIsAndBChange.CheckedChanged += new System.EventHandler(this.cbIsAtoB相抵_CheckedChanged);
                    this.intMiningSection.ValueChanged += new System.EventHandler(this.intMiningSection_ValueChanged);
                    this.intDisciplinaryMeasures.ValueChanged += new System.EventHandler(this.intDisciplinaryMeasures_ValueChanged);
                    this.cbIsePaper.CheckedChanged += new System.EventHandler(this.cbIsePaper_CheckedChanged);
                }
                else
                {
                    MsgBox.Show("作業已終止");
                }
            }
            else
            {
                MsgBox.Show("發生錯誤:\n" + e.Error.Message);

            }
        }


        private void btPrint_Click(object sender, EventArgs e)
        {
            if (!bgw.IsBusy)
            {
                btPrint.Enabled = false;
                _SchoolYear = integerInput1.Text;
                _Semester = integerInput2.Text;
                bgw.RunWorkerAsync();
            }
            else
            {
                MsgBox.Show("目前忙碌中");
            }
        }

        private void Bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            //儲存設定
            ConfigData cd = K12.Data.School.Configuration[ConfigName2];
            cd[ConfigName2_1] = "" + _MiningSection;
            cd[ConfigName2_2] = "" + _DisciplinaryMeasures;
            cd[ConfigName2_換算] = "" + _IsAtoB換算;
            cd[ConfigName2_相抵] = "" + _IsAtoB相抵;
            cd[ConfigName2_電子報] = "" + _IsePaper電子報;
            cd[ConfigName2_預設推播] = _MessageName;
            cd.Save();

            SendMessageTeacher = new List<string>();

            bgw.ReportProgress(5, "取得班級學生");
            //取得班級學生
            List<string> ClassIDList = K12.Presentation.NLDPanels.Class.SelectedSource;
            Dictionary<string, StudRecord> StudDic = GetClassStudent(ClassIDList);

            //2020/5/4 - 取得班級清單
            //班級ID 班級學生ID
            Dictionary<string, List<StudRecord>> ClassIDDic = GetClassList(StudDic);

            bool 沒有符合資料 = true;

            if (StudDic.Count > 0)
            {
                bgw.ReportProgress(15, "取得缺曠資料");
                //缺曠資料
                GetAttendance(StudDic);

                //獎懲資料
                bgw.ReportProgress(30, "取得獎懲資料");
                GetMeDemrit(StudDic);

                ColumnInTitleIndex = tool.GetColumnTitle(_IsAtoB換算, _IsAtoB相抵);

                wb = new Workbook(new MemoryStream(Properties.Resources.日常生活預警表));
                wb_clsss = new Workbook();
                wb_clsss.Worksheets.Clear();

                sy = wb.CreateStyle();
                sy.Font.Size = 10; //字型
                sy.IsTextWrapped = true; //自動換行
                sy.HorizontalAlignment = TextAlignmentType.Center; //水平置中
                sy.VerticalAlignment = TextAlignmentType.Center; //垂直置中
                sy.Font.Name = "標楷體";

                //取得缺曠設定
                AbsenceConfig1 = LoadAbsence(ConfigName1);
                AbsenceConfig2 = LoadAbsence(ConfigName3);

                bgw.ReportProgress(45, "建立缺曠格式");
                #region 建立缺曠格式

                byte AttendanceColumn = 0;

                //歷年功過換算
                if (_IsAtoB換算)
                {
                    #region 建立功過換算格式

                    byte MeDemritColumn = 19;


                    int tempMeD = MeDemritColumn;
                    int 總共多少功過換算 = 0;
                    for (int x = 0; x < 6; x++)
                    {
                        wb.Worksheets[SheetIndex].Cells.InsertColumn(tempMeD);
                        wb.Worksheets["範本"].Cells.InsertColumn(tempMeD);
                        總共多少功過換算++;
                    }

                    int MergeMeDIndex = MeDemritColumn;
                    wb.Worksheets[SheetIndex].Cells.Merge(0, 0, 1, MergeMeDIndex + 總共多少功過換算 + 1); //標題
                    wb.Worksheets["範本"].Cells.Merge(0, 0, 1, MergeMeDIndex + 總共多少功過換算 + 1); //標題

                    wb.Worksheets["範本"].Cells.Merge(1, MergeMeDIndex, 1, 總共多少功過換算); //功過相抵標題

                    #endregion

                    #region 填入功過換算標題

                    wb.Worksheets[SheetIndex].Cells[1, MeDemritColumn].PutValue("歷年功過換算");
                    wb.Worksheets["範本"].Cells[1, MeDemritColumn].PutValue("歷年功過換算");

                    SeetTitle("換算大功", "大功", MeDemritColumn);
                    SeetTitle("換算小功", "小功", MeDemritColumn += 1);
                    SeetTitle("換算嘉獎", "嘉獎", MeDemritColumn += 1);
                    SeetTitle("換算大過", "大過", MeDemritColumn += 1);
                    SeetTitle("換算小過", "小過", MeDemritColumn += 1);
                    SeetTitle("換算警告", "警告", MeDemritColumn += 1);

                    #endregion

                    AttendanceColumn = 25;

                }
                else
                {
                    AttendanceColumn = 19;
                }

                if (_IsAtoB相抵)
                {
                    #region 建立功過相抵格式

                    byte MeDemritColumn = 0;
                    if (_IsAtoB換算)
                    {
                        AttendanceColumn = 31;
                        MeDemritColumn = 25;
                    }
                    else
                    {
                        AttendanceColumn = 25;
                        MeDemritColumn = 19;
                    }

                    int tempMeD = MeDemritColumn;
                    int 總共多少功過相抵 = 0;
                    for (int x = 0; x < 6; x++)
                    {
                        wb.Worksheets[SheetIndex].Cells.InsertColumn(tempMeD);
                        wb.Worksheets["範本"].Cells.InsertColumn(tempMeD);
                        總共多少功過相抵++;
                    }

                    int MergeMeDIndex = MeDemritColumn;
                    wb.Worksheets[SheetIndex].Cells.Merge(0, 0, 1, MergeMeDIndex + 總共多少功過相抵 + 1); //標題
                    wb.Worksheets["範本"].Cells.Merge(0, 0, 1, MergeMeDIndex + 總共多少功過相抵 + 1); //標題

                    wb.Worksheets["範本"].Cells.Merge(1, MergeMeDIndex, 1, 總共多少功過相抵); //功過相抵標題

                    #endregion

                    #region 填入功過相抵標題

                    wb.Worksheets[SheetIndex].Cells[1, MeDemritColumn].PutValue("歷年功過相抵");
                    wb.Worksheets["範本"].Cells[1, MeDemritColumn].PutValue("歷年功過相抵");

                    SeetTitle("相抵大功", "大功", MeDemritColumn);
                    SeetTitle("相抵小功", "小功", MeDemritColumn += 1);
                    SeetTitle("相抵嘉獎", "嘉獎", MeDemritColumn += 1);
                    SeetTitle("相抵大過", "大過", MeDemritColumn += 1);
                    SeetTitle("相抵小過", "小過", MeDemritColumn += 1);
                    SeetTitle("相抵警告", "警告", MeDemritColumn += 1);

                    #endregion

                }
                else
                {
                    if (_IsAtoB換算)
                        AttendanceColumn = 25;
                    else
                        AttendanceColumn = 19;

                }

                int temp = AttendanceColumn;

                int 總共多少 = 0;
                foreach (string each in AbsenceConfig1.Keys)
                {
                    int countColumn = 0;
                    foreach (string eachIn in AbsenceConfig1[each])
                    {
                        wb.Worksheets[SheetIndex].Cells.InsertColumn(temp);
                        wb.Worksheets["範本"].Cells.InsertColumn(temp);
                        countColumn++;
                        總共多少++;
                    }
                }

                int MergeIndex = AttendanceColumn;
                wb.Worksheets[SheetIndex].Cells.Merge(0, 0, 1, MergeIndex + 總共多少 + 1); //標題
                wb.Worksheets["範本"].Cells.Merge(0, 0, 1, MergeIndex + 總共多少 + 1); //標題
                if (總共多少 != 0)
                {
                    wb.Worksheets["範本"].Cells.Merge(1, MergeIndex, 1, 總共多少); //缺曠標題
                    wb.Worksheets[SheetIndex].Cells[1, AttendanceColumn].PutValue("本期缺曠");
                    wb.Worksheets["範本"].Cells[1, AttendanceColumn].PutValue("本期缺曠");
                }

                foreach (string each in AbsenceConfig1.Keys)
                {
                    foreach (string eachIn in AbsenceConfig1[each])
                    {
                        string name = each + eachIn;
                        ColumnInTitleIndex.Add(name, AttendanceColumn);
                        SeetTitle(each + eachIn, each + eachIn, AttendanceColumn);
                        AttendanceColumn++;
                    }
                }

                #endregion

                #region 其他欄位

                ColumnInTitleIndex.Add("是否留查", AttendanceColumn);
                ColumnInTitleIndex.Add("是否全勤", AttendanceColumn += 1);
                ColumnInTitleIndex.Add("備註", AttendanceColumn += 1);

                //wb.Worksheets["範本"].Cells.CreateRange(4, 0, 1, AttendanceColumn - 1).SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                //wb.Worksheets["範本"].Cells.CreateRange(4, 0, 1, AttendanceColumn - 1).SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                //wb.Worksheets["範本"].Cells.CreateRange(4, 0, 1, AttendanceColumn - 1).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                //wb.Worksheets["範本"].Cells.CreateRange(4, 0, 1, AttendanceColumn - 1).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                //Style style_studentRow = wb.Worksheets["範本"].Cells[4, 1].GetStyle();
                //wb.Worksheets["範本"].Cells.CreateRange(4, 0, 1, AttendanceColumn - 1).SetStyle(style_studentRow);

                #endregion

                RangeDetil = wb.Worksheets["範本"].Cells.CreateRange(0, 3, false);
                RangeRow = wb.Worksheets["範本"].Cells.CreateRange(3, 1, false);

                int rowIndex = 0; //
                wb.Worksheets[SheetIndex].Cells.CreateRange(rowIndex, 3, false).CopyStyle(RangeDetil);
                wb.Worksheets[SheetIndex].Cells.CreateRange(rowIndex, 3, false).CopyValue(RangeDetil);

                StringBuilder sb = new StringBuilder();
                sb.Append(School.ChineseName + "　");
                sb.Append(_SchoolYear + "學年度　");
                sb.Append("第" + _Semester + "學期　");
                sb.Append("日常生活表現預警表");

                wb.Worksheets[SheetIndex].Cells[rowIndex, 0].PutValue(sb.ToString());
                wb.Worksheets["範本"].Cells[rowIndex, 0].PutValue(sb.ToString());

                rowIndex += 3;
                //班級
                List<string> TeacherIDNameList = new List<string>();
                bgw.ReportProgress(70, "整理獎懲/缺曠/功過換算");
                foreach (string classID in ClassIDDic.Keys)
                {
                    int newRowIndex = 3;

                    //每一個班級之後
                    //增加一篇標頭
                    List<StudRecord> studList = ClassIDDic[classID];
                    studList.Sort(SortStudent);

                    //每一個Sheet 班級名稱{12345} <-系統編號
                    string name = studList[0].ClassName + "{" + studList[0].TeacherID + "}";

                    //沒有教師,無法發送電子報表
                    if (studList[0].TeacherID == "")
                        continue;

                    //教師系統編號
                    if (TeacherIDNameList.Contains(name))
                        continue;

                    int class_a = wb_clsss.Worksheets.Add();
                    wb_clsss.Worksheets[class_a].Copy(wb.Worksheets["範本"]);
                    wb_clsss.Worksheets[class_a].Name = name;
                    wb_clsss.Worksheets[class_a].AutoFitRows();

                    if (!TeacherIDNameList.Contains(name))
                        TeacherIDNameList.Add(name);

                    //學生
                    int pointStudCount = 0;
                    foreach (StudRecord SR in studList)
                    {

                        //整理獎懲
                        SR.RunDeMerit(_DisciplinaryMeasures);

                        //整理缺曠
                        SR.RunAttendance(periodDic, AbsenceConfig1, AbsenceConfig2);

                        if (_IsAtoB換算)
                        {
                            SR.RunOffset();
                        }

                        if (_IsAtoB相抵)
                        {
                            SR.RunOffsetAtoB();

                        }

                        if (SR.CheckSection(_MiningSection, _DisciplinaryMeasures, _IsAtoB相抵, _IsAtoB換算))
                        {
                            pointStudCount++;
                            //如果這名學生列印資料
                            //則通知老師
                            if (!SendMessageTeacher.Contains(SR.TeacherID))
                            {
                                SendMessageTeacher.Add(SR.TeacherID);
                            }

                            //主頁面
                            SetWorksheets(wb, SR, SheetIndex, rowIndex);
                            rowIndex++;
                            //電子報表單頁
                            SetWorksheets(wb_clsss, SR, class_a, newRowIndex);
                            newRowIndex++;

                            沒有符合資料 = false;
                        }
                        else
                        {
                            //沒有符合資料
                        }
                    }

                    //不予列印時,將報表從wb_clsss移除
                    if (pointStudCount == 0)
                        wb_clsss.Worksheets.RemoveAt(class_a);
                }
                bgw.ReportProgress(100, "儲存");
                wb.Worksheets.RemoveAt("範本"); //原本範本是用來複製標題
            }
            else
            {
                e.Cancel = true;
            }

            if (沒有符合資料)
            {
                e.Cancel = true;
            }
        }

        private void SetWorksheets(Workbook book, StudRecord SR, int sheetIndex, int rowIndex)
        {
            book.Worksheets[sheetIndex].Cells.CreateRange(rowIndex, 1, false).CopyStyle(RangeRow);
            book.Worksheets[sheetIndex].Cells.CreateRange(rowIndex, 1, false).CopyValue(RangeRow);

            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["班級"]].PutValue(SR.ClassName);
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["學號"]].PutValue(SR.StudentNumber);
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["座號"]].PutValue(SR.SeatNo);
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["姓名"]].PutValue(SR.StudentName);

            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["本期大功"]].PutValue(CByZero(SR.本期大功));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["本期小功"]].PutValue(CByZero(SR.本期小功));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["本期嘉獎"]].PutValue(CByZero(SR.本期嘉獎));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["本期大過"]].PutValue(CByZero(SR.本期大過));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["本期小過"]].PutValue(CByZero(SR.本期小過));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["本期警告"]].PutValue(CByZero(SR.本期警告));

            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年大功"]].PutValue(CByZero(SR.歷年大功));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年小功"]].PutValue(CByZero(SR.歷年小功));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年嘉獎"]].PutValue(CByZero(SR.歷年嘉獎));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年大過"]].PutValue(CByZero(SR.歷年大過));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年小過"]].PutValue(CByZero(SR.歷年小過));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年警告"]].PutValue(CByZero(SR.歷年警告));

            if (_IsAtoB換算)
            {
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過換算大功"]].PutValue(CByZero(SR.歷年功過換算大功));
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過換算小功"]].PutValue(CByZero(SR.歷年功過換算小功));
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過換算嘉獎"]].PutValue(CByZero(SR.歷年功過換算嘉獎));
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過換算大過"]].PutValue(CByZero(SR.歷年功過換算大過));
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過換算小過"]].PutValue(CByZero(SR.歷年功過換算小過));
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過換算警告"]].PutValue(CByZero(SR.歷年功過換算警告));
            }

            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年銷過大過"]].PutValue(CByZero(SR.歷年銷過大過));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年銷過小過"]].PutValue(CByZero(SR.歷年銷過小過));
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年銷過警告"]].PutValue(CByZero(SR.歷年銷過警告));

            if (_IsAtoB相抵)
            {
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過相抵大功"]].PutValue(CByZero(SR.歷年功過相抵大功));
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過相抵小功"]].PutValue(CByZero(SR.歷年功過相抵小功));
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過相抵嘉獎"]].PutValue(CByZero(SR.歷年功過相抵嘉獎));
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過相抵大過"]].PutValue(CByZero(SR.歷年功過相抵大過));
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過相抵小過"]].PutValue(CByZero(SR.歷年功過相抵小過));
                book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["歷年功過相抵警告"]].PutValue(CByZero(SR.歷年功過相抵警告));
            }

            foreach (string att in SR.AttendanceDic.Keys)
            {
                if (ColumnInTitleIndex.ContainsKey(att))
                {
                    book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex[att]].PutValue(SR.AttendanceDic[att]);
                }
            }

            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["是否留查"]].PutValue(SR.是否留查 ? "是" : "");
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["是否全勤"]].PutValue(SR.是否全勤 ? "是" : "");

            //本筆資料的確定源
            book.Worksheets[sheetIndex].Cells[rowIndex, ColumnInTitleIndex["備註"]].PutValue(string.Join(",\n", SR.RemarkList));
        }

        private void Bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btPrint.Enabled = true;
            this.Text = "日常生活表現預警表";

            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            if (!e.Cancelled)
            {
                #region 背景工作完成後...

                FISCA.Presentation.MotherForm.SetStatusBarMessage("日常生活表現預警表,列印完成!!");

                //Save To ePaper
                //if (_IsePaper電子報)
                //{
                //    MemoryStream memoryStream = new MemoryStream();
                //    wb_clsss.Save(memoryStream, SaveFormat.Xlsx);
                //    ePaperCloud.upload_ePaper(int.Parse(_SchoolYear), int.Parse(_Semester), string.Format("日常生活表現預警表_{0}學年度_第{1}學期", _SchoolYear, _Semester), "", memoryStream, ePaperCloud.ViewerType.Teacher, ePaperCloud.FormatType.Xlsx, string.Format(_MessageName, _SchoolYear, _Semester));
                //}
                //else
                //{
                SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = string.Format("日常生活表現預警表_{0}學年度_第{1}學期.xlsx", _SchoolYear, _Semester);
                sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        wb.Save(sd.FileName);
                        System.Diagnostics.Process.Start(sd.FileName);

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        this.Enabled = true;
                        return;
                    }
                }
                //}
                #endregion
            }
            else
            {
                MsgBox.Show("列印取消,查無資料");
            }

        }

        private void SeetTitle(string _deDemritValue, string name, int _deDemritColumn)
        {
            wb.Worksheets[SheetIndex].Cells[2, _deDemritColumn].SetStyle(sy);
            wb.Worksheets[SheetIndex].Cells[2, _deDemritColumn].PutValue(name);
            wb.Worksheets[SheetIndex].Cells.SetColumnWidth(_deDemritColumn, 2); //column寬度
            wb.Worksheets[SheetIndex].Cells.SetRowHeight(2, 60); //Row高度

            wb.Worksheets["範本"].Cells[2, _deDemritColumn].SetStyle(sy);
            wb.Worksheets["範本"].Cells[2, _deDemritColumn].PutValue(name);
            wb.Worksheets["範本"].Cells.SetColumnWidth(_deDemritColumn, 2); //column寬度
            wb.Worksheets["範本"].Cells.SetRowHeight(2, 60); //Row高度

            wb.Worksheets[SheetIndex].Cells.CreateRange(2, _deDemritColumn, 1, 1).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            wb.Worksheets["範本"].Cells.CreateRange(2, _deDemritColumn, 1, 1).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

        }

        private int SortStudent(StudRecord x, StudRecord y)
        {
            //年級
            string studentA = x.GradeYear.PadLeft(3, '0');
            string studentB = y.GradeYear.PadLeft(3, '0');

            //排序
            studentA += x.ClassNumber.PadLeft(3, '0');
            studentB += y.ClassNumber.PadLeft(3, '0');

            //名稱
            studentA += x.ClassName.PadLeft(10, '0');
            studentB += y.ClassName.PadLeft(10, '0');

            studentA += x.SeatNo.PadLeft(2, '0');
            studentB += y.SeatNo.PadLeft(2, '0');

            studentA += x.StudentName.PadLeft(10, '0');
            studentB += y.StudentName.PadLeft(10, '0');

            return studentA.CompareTo(studentB);
        }

        /// <summary>
        /// 如果是0,回傳空字串
        /// </summary>
        private string CByZero(int x)
        {
            if (x == 0)
                return "";
            else
                return x.ToString();
        }


        /// <summary>
        /// 由班級系統編號
        /// 取得學生清單
        /// </summary>
        private Dictionary<string, StudRecord> GetClassStudent(List<string> classIDList)
        {
            //功過換算表
            reduceRecord = K12.Data.MeritDemeritReduce.Select();

            //由班級ID取得學生系統編號與資料
            DataTable dt = tool._Q.Select(string.Format(@"select class.id as class_id,class.class_name,
student.id as student_id,student.name as student_name,student.student_number,
student.seat_no,class.display_order,class.grade_year,teacher.id as teacher_id 
from class 
join student on student.ref_class_id=class.id
left join teacher on class.ref_teacher_id=teacher.id 
where class.id in ({0}) and student.status in (1 , 2) 
order by class.grade_year,class.display_order,class.class_name,student.seat_no,student.name", string.Join(",", classIDList))) ;

            Dictionary<string, StudRecord> dic = new Dictionary<string, StudRecord>();
            foreach (DataRow row in dt.Rows)
            {
                string student_id = "" + row["student_id"];
                if (!dic.ContainsKey(student_id))
                {
                    dic.Add(student_id, new StudRecord(row, _SchoolYear, _Semester, reduceRecord));
                }
            }

            return dic;
        }

        /// <summary>
        /// 由班級系統編號
        /// 取得學生清單
        /// </summary>
        private Dictionary<string, List<StudRecord>> GetClassList(Dictionary<string, StudRecord> studDic)
        {
            //班級ID 學生ID 學生
            Dictionary<string, List<StudRecord>> dic = new Dictionary<string, List<StudRecord>>();

            foreach (StudRecord each in studDic.Values)
            {
                if (!dic.ContainsKey(each.ClassID))
                {
                    dic.Add(each.ClassID, new List<StudRecord>());
                }
                dic[each.ClassID].Add(each);
            }

            return dic;
        }

        private void GetAttendance(Dictionary<string, StudRecord> studDic)
        {
            //取得學生所有的缺曠紀錄
            List<AttendanceRecord> AttendanceList = Attendance.SelectByStudentIDs(studDic.Keys);

            foreach (AttendanceRecord each in AttendanceList)
            {
                if (studDic.ContainsKey(each.RefStudentID))
                {
                    studDic[each.RefStudentID].AttendanceList.Add(each);
                }
            }
        }

        private void GetMeDemrit(Dictionary<string, StudRecord> studDic)
        {
            //取得學生所有的獎懲紀錄
            List<MeritRecord> meriteList = Merit.SelectByStudentIDs(studDic.Keys);
            foreach (MeritRecord each in meriteList)
            {
                if (studDic.ContainsKey(each.RefStudentID))
                {
                    studDic[each.RefStudentID].MeritList.Add(each);
                }
            }

            List<DemeritRecord> demeriteList = Demerit.SelectByStudentIDs(studDic.Keys);
            foreach (DemeritRecord each in demeriteList)
            {
                if (studDic.ContainsKey(each.RefStudentID))
                {
                    studDic[each.RefStudentID].DemeritList.Add(each);

                    //懲戒資料內包含留查資料

                }
            }
        }

        /// <summary>
        /// 畫面開啟或關閉
        /// </summary>
        bool SetloadForm
        {
            set
            {
                integerInput1.Enabled = value;
                integerInput2.Enabled = value;
                intMiningSection.Enabled = value;
                intDisciplinaryMeasures.Enabled = value;
                cbAtoBSum.Enabled = value;
                cbIsAndBChange.Enabled = value;
                cbIsePaper.Enabled = value;
                btPrint.Enabled = value;
            }
        }

        private void linkSetupAttendance_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SelectTypeForm form = new SelectTypeForm(ConfigName1);
            form.ShowDialog();
        }

        private void linkMiningSection_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SelectTypeForm form = new SelectTypeForm(ConfigName3);
            form.ShowDialog();
        }

        private void intMiningSection_ValueChanged(object sender, EventArgs e)
        {
            _MiningSection = intMiningSection.Value;
        }

        private void intDisciplinaryMeasures_ValueChanged(object sender, EventArgs e)
        {
            _DisciplinaryMeasures = intDisciplinaryMeasures.Value;
        }

        private void cbAtoB換算_CheckedChanged(object sender, EventArgs e)
        {
            this.cbAtoBSum.CheckedChanged -= new System.EventHandler(this.cbAtoB換算_CheckedChanged);
            this.cbIsAndBChange.CheckedChanged -= new System.EventHandler(this.cbIsAtoB相抵_CheckedChanged);

            ChangeLabel();
            
            _IsAtoB換算 = cbAtoBSum.Checked;

            this.cbAtoBSum.CheckedChanged += new System.EventHandler(this.cbAtoB換算_CheckedChanged);
            this.cbIsAndBChange.CheckedChanged += new System.EventHandler(this.cbIsAtoB相抵_CheckedChanged);
        }

        private void cbIsAtoB相抵_CheckedChanged(object sender, EventArgs e)
        {
            //功過相抵必先勾選功過換算
            this.cbAtoBSum.CheckedChanged -= new System.EventHandler(this.cbAtoB換算_CheckedChanged);
            this.cbIsAndBChange.CheckedChanged -= new System.EventHandler(this.cbIsAtoB相抵_CheckedChanged);

            cbAtoBSum.Enabled = !cbIsAndBChange.Checked;
            cbAtoBSum.Checked = cbIsAndBChange.Checked ? true : false;

            ChangeLabel();

            _IsAtoB換算 = cbAtoBSum.Checked;
            _IsAtoB相抵 = cbIsAndBChange.Checked;

            this.cbAtoBSum.CheckedChanged += new System.EventHandler(this.cbAtoB換算_CheckedChanged);
            this.cbIsAndBChange.CheckedChanged += new System.EventHandler(this.cbIsAtoB相抵_CheckedChanged);
        }

        private void ChangeLabel()
        {
            if (cbIsAndBChange.Checked)
            {
                labelX1.Text = "*.懲戒採計[歷年功過相抵]結果";
                labelX1.ForeColor = Color.Red;
            }
            else
            {
                if (cbAtoBSum.Checked)
                {
                    labelX1.Text = "*.懲戒採計[歷年功過換算]結果";
                    labelX1.ForeColor = Color.Red;
                }
                else
                {
                    labelX1.Text = "*.懲戒採計[歷年獎懲]結果";
                    labelX1.ForeColor = Color.Gray;
                }
            }
        }

        private void cbIsePaper_CheckedChanged(object sender, EventArgs e)
        {
            _IsePaper電子報 = cbIsePaper.Checked;
        }

        private void linkViewContent_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ViewMessage vmessage = new ViewMessage(_MessageName);
            DialogResult dr = vmessage.ShowDialog();

            if (dr == DialogResult.Yes)
            {
                _MessageName = vmessage._MessageName;
            }
        }

        private void btExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
