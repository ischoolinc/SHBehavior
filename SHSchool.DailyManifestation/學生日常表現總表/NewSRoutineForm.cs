using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.Presentation;
using K12.Data;
using System.IO;
using System.Diagnostics;
using System.Xml;
using Campus.Configuration;
using Aspose.Words;
using SmartSchool.ePaper;

namespace SHSchool.DailyManifestation
{
    public partial class NewSRoutineForm : BaseForm
    {

        private string ConfigName = "SHSchool.DailyManifestation.日常生活記錄表";

        private BackgroundWorker BGW = new BackgroundWorker();
        //主文件
        private Document _doc;
        //單頁範本
        private Document _template;
        //移動使用
        private Run _run;

        List<string> DLBList1 = new List<string>();
        List<string> DLBList2 = new List<string>();

        List<string> BehaviorList;
        List<string> BehaviorList1 = new List<string>();
        List<string> BehaviorList2 = new List<string>();
        List<string> BehaviorList3 = new List<string>();

        Dictionary<string, int> DicSummaryIndex = new Dictionary<string, int>();
        Dictionary<string, string> UpdateCoddic = new Dictionary<string, string>();

        int SuperBehaviorIndex1 = 5;
        int SuperBehaviorIndex2 = 13;

        bool PrintUpdateStudentFile = false;

        /// <summary>
        /// 學生電子報表
        /// </summary>
        SmartSchool.ePaper.ElectronicPaper paperForStudent { get; set; }

        List<string> absenceList;

        bool SingeFile = false;

        //學生資料集
        StudentInfo Data;

        Campus.Configuration.ConfigData cd;

        /// <summary>
        /// 記錄日常生活表現 - 學期的第一格
        /// </summary>
        Dictionary<string, Cell> DicBack;

        public NewSRoutineForm()
        {
            InitializeComponent();
        }

        private void NewSRoutineForm_Load(object sender, EventArgs e)
        {
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            MotherForm.SetStatusBarMessage("開始列印學生日常表現總表...");
            btnSave.Enabled = false;

            SetNameIndex();

            SingeFile = cbSingeFile.Checked;
            PrintUpdateStudentFile = checkBoxX1.Checked;

            BGW.RunWorkerAsync();
        }

        /// <summary>
        /// 背景模式
        /// </summary>
        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
             paperForStudent = new SmartSchool.ePaper.ElectronicPaper(School.DefaultSchoolYear + "學生日常表現總表", School.DefaultSchoolYear, School.DefaultSemester, SmartSchool.ePaper.ViewerType.Student);

            #region 設定檔

            cd = Campus.Configuration.Config.User[ConfigName];

            XmlElement config = cd.GetXml("XmlData", null);
            if (config == null)
            {
                e.Cancel = true;
                return;
            }

            absenceList = new List<string>();
            int absenceindex = 1;
            //集會,一般...
            foreach (XmlElement xml in config.SelectNodes("Type"))
            {
                //曠課,病假...
                foreach (XmlElement node in xml.SelectNodes("Absence"))
                {
                    string s1 = xml.GetAttribute("Text");
                    string s2 = node.GetAttribute("Text");
                    absenceList.Add(s2 + s1);

                    //記錄Index
                    DicSummaryIndex.Add(s2 + s1, absenceindex);
                    absenceindex++;
                }
            }

            #endregion

            #region 開頭處理

            BehaviorList = GetBehaviorConfig();
            BehaviorList1 = new List<string>();
            BehaviorList2 = new List<string>();
            BehaviorList3 = new List<string>();

            //製作日常生活表現,標題的Index
            int BeIndex = 0;
            foreach (string each in BehaviorList)
            {
                BeIndex++;
                if (BeIndex > SuperBehaviorIndex1 && BeIndex <= SuperBehaviorIndex2)
                {
                    BehaviorList2.Add(each);
                    DicSummaryIndex.Add(each, BeIndex - SuperBehaviorIndex1);
                }
                else if (BeIndex > SuperBehaviorIndex2)
                {
                    BehaviorList3.Add(each);
                    DicSummaryIndex.Add(each, BeIndex - SuperBehaviorIndex2);
                }
                else if (BeIndex <= SuperBehaviorIndex1)
                {
                    BehaviorList1.Add(each);
                    DicSummaryIndex.Add(each, BeIndex);
                }
            }

            //依據日常生活表現欄位,讀取範本
            _template = new Document(new MemoryStream(Properties.Resources.日常生活記錄表_範本_1));

            if (BehaviorList.Count > SuperBehaviorIndex1)
                _template = new Document(new MemoryStream(Properties.Resources.日常生活記錄表_範本_2));

            if (BehaviorList.Count > SuperBehaviorIndex2)
                _template = new Document(new MemoryStream(Properties.Resources.日常生活記錄表_範本_3));

            //建立資料模型
            Data = new StudentInfo(BehaviorList);

            #endregion

            if (SingeFile) //如果是單檔儲存
            {
                #region 單一學生

                Dictionary<string, Document> DocDic = new Dictionary<string, Document>();

                foreach (string student in Data.DicStudent.Keys)
                {
                    Document PageOne = SetDocument(student);

                    //電子報表
                    MemoryStream stream = new MemoryStream();
                    PageOne.Save(stream, SaveFormat.Doc);
                    paperForStudent.Append(new PaperItem(PaperFormat.Office2003Doc, stream, student));

                    if (!DocDic.ContainsKey(student))
                    {
                        DocDic.Add(student, PageOne);
                    }
                }

                e.Result = DocDic;

                #endregion
            }
            else
            {
                #region 多名學生存單一檔案

                _doc = new Document();
                _doc.Sections.Clear(); //清空此Document

                foreach (string student in Data.DicStudent.Keys)
                {
                    Document PageOne = SetDocument(student);

                    //電子報表
                    MemoryStream stream = new MemoryStream();
                    PageOne.Save(stream, SaveFormat.Doc);
                    paperForStudent.Append(new PaperItem(PaperFormat.Office2003Doc, stream, student));

                    _doc.Sections.Add(_doc.ImportNode(PageOne.FirstSection, true));
                }

                e.Result = _doc;

                #endregion
            }

            //如果有打勾則上傳電子報表
            if (PrintUpdateStudentFile)
                 SmartSchool.ePaper.DispatcherProvider.Dispatch(paperForStudent);
        }

        /// <summary>
        /// 每一名學生的報表資料列印
        /// </summary>
        /// <returns></returns>
        private Document SetDocument(string student)
        {
            #region 每一名學生的報表資料列印

            StudentDataObj obj = Data.DicStudent[student];
            //取得範本樣式
            Document PageOne = (Document)_template.Clone(true);
            //???
            _run = new Run(PageOne);
            //可建構的...
            DocumentBuilder builder = new DocumentBuilder(PageOne);
            DocumentBuilder builderX = new DocumentBuilder(_template);

            builder.MoveToMergeField("缺曠");
            Cell absenceCell = (Cell)builder.CurrentParagraph.ParentNode;
            foreach (string each in absenceList)
            {
                Write(absenceCell, each);

                if (absenceCell.NextSibling != null)
                {
                    absenceCell = absenceCell.NextSibling as Cell; //取得下一格
                }
            }

            #region 資料MailMerge第一步

            List<string> name = new List<string>();
            List<string> value = new List<string>();

            name.Add("學校名稱");
            value.Add(School.ChineseName);

            name.Add("學號");
            value.Add(obj.StudentRecord.StudentNumber);

            name.Add("姓名");
            value.Add(obj.StudentRecord.Name);

            name.Add("性別");
            value.Add(obj.StudentRecord.Gender);

            name.Add("身分證");
            value.Add(obj.StudentRecord.IDNumber);

            name.Add("生日");
            value.Add(obj.StudentRecord.Birthday.HasValue ? obj.StudentRecord.Birthday.Value.ToShortDateString() : "");

            name.Add("出生");
            value.Add(obj.StudentRecord.BirthPlace);

            name.Add("監護");
            value.Add(obj.CustodianName);

            name.Add("戶籍");
            value.Add(obj.AddressPermanent);

            name.Add("電話1");
            value.Add(obj.PhonePermanent);

            name.Add("稱謂");
            value.Add(obj.CustodianTitle);

            name.Add("通訊");
            value.Add(obj.AddressMailing);

            name.Add("電話2");
            value.Add(obj.PhoneContact);

            name.Add("一上");
            value.Add(obj.GradeYear11);

            name.Add("一下");
            value.Add(obj.GradeYear12);

            name.Add("二上");
            value.Add(obj.GradeYear21);

            name.Add("二下");
            value.Add(obj.GradeYear22);

            name.Add("三上");
            value.Add(obj.GradeYear31);

            name.Add("三下");
            value.Add(obj.GradeYear32);

            PageOne.MailMerge.Execute(name.ToArray(), value.ToArray());
            #endregion

            #region 依據日常生活表現數量,建立欄位數量

            //記錄日常生活表現 - 學期的第一格
            DicBack = new Dictionary<string, Cell>();

            builder.MoveToMergeField("異2");
            SetRowCount("異動資料欄1", (Cell)builder.CurrentParagraph.ParentNode, obj.ListUpdateRecord.Count);

            builder.MoveToMergeField("日2");
            SetRowCount("日常生活表現欄1", (Cell)builder.CurrentParagraph.ParentNode, obj.TextScoreDic.Count);


            if (BehaviorList2.Count > 0)
            {
                builder.MoveToMergeField("常2");
                SetRowCount("日常生活表現欄2", (Cell)builder.CurrentParagraph.ParentNode, obj.TextScoreDic.Count);
            }

            if (BehaviorList3.Count > 0)
            {
                builder.MoveToMergeField("生2");
                SetRowCount("日常生活表現欄3", (Cell)builder.CurrentParagraph.ParentNode, obj.TextScoreDic.Count);
            }

            builder.MoveToMergeField("統2");
            SetRowCount("缺曠資料欄1", (Cell)builder.CurrentParagraph.ParentNode, obj.DicAsbs.Count);

            builder.MoveToMergeField("獎2");
            SetRowCount("獎懲資料欄1", (Cell)builder.CurrentParagraph.ParentNode, obj.DicMeritDemerit.Count);

            #endregion

            #region 異動處理

            Cell UpdateRecordCell = DicBack["異動資料欄1"];

            foreach (UpdateRecordRecord updateRecord in obj.ListUpdateRecord)
            {
                List<string> list = new List<string>();
                list.Add(updateRecord.SchoolYear.HasValue ? updateRecord.SchoolYear.Value.ToString() : "");
                list.Add(updateRecord.Semester.HasValue ? updateRecord.Semester.Value.ToString() : "");
                list.Add(updateRecord.UpdateDate);
                list.Add(updateRecord.ADDate);
                list.Add(updateRecord.UpdateDescription);
                //list.Add(GetUpdateRecordCode(updateRecord.UpdateCode));
                list.Add(updateRecord.ADNumber);
                list.Add(updateRecord.Comment);

                foreach (string UpdateName in list)
                {
                    //寫入
                    Write(UpdateRecordCell, UpdateName);
                    if (UpdateRecordCell.NextSibling != null) //是否最後一格
                        UpdateRecordCell = UpdateRecordCell.NextSibling as Cell; //下一格
                }
                Row Nextrow = UpdateRecordCell.ParentRow.NextSibling as Row; //取得下一個Row
                UpdateRecordCell = Nextrow.FirstCell; //第一格Cell   

                if (UpdateRecordCell == null)
                    break;
            }

            #endregion

            #region 日1

            //填入日常生活表現1標題內容
            builder.MoveToMergeField("日1");
            Cell setupCell = (Cell)builder.CurrentParagraph.ParentNode;

            foreach (string each in BehaviorList1)
            {
                Write(setupCell, each); //填入導師評語

                if (setupCell.NextSibling != null)
                {
                    setupCell = setupCell.NextSibling as Cell; //取得下一格
                }
            }

            Cell MoralScore1 = DicBack["日常生活表現欄1"];
            //KEY 學年期資訊
            foreach (string moralScore in obj.TextScoreDic.Keys)
            {
                //填入學年期
                Write(MoralScore1, moralScore);
                //欄位名稱
                foreach (string BehaviorConfigName1 in obj.TextScoreDic[moralScore].Keys)
                {
                    //第一欄內容
                    if (BehaviorList1.Contains(BehaviorConfigName1))
                    {
                        Cell Score = GetMoveRightCell(MoralScore1, GetSummaryIndex(BehaviorConfigName1));
                        Write(Score, obj.TextScoreDic[moralScore][BehaviorConfigName1]);
                    }
                }

                Row Nextrow = MoralScore1.ParentRow.NextSibling as Row; //取得下一行
                if (Nextrow == null)
                    break;
                MoralScore1 = Nextrow.FirstCell; //第一格
            }

            #endregion

            #region 日2

            if (BehaviorList2.Count > 0)
            {
                //填入日常生活表現2標題內容
                builder.MoveToMergeField("常1");
                setupCell = (Cell)builder.CurrentParagraph.ParentNode;

                foreach (string each in BehaviorList2)
                {
                    Write(setupCell, each); //填入導師評語

                    if (setupCell.NextSibling != null)
                    {
                        setupCell = setupCell.NextSibling as Cell; //取得下一格
                    }
                }

                Cell MoralScore2 = DicBack["日常生活表現欄2"];

                //KEY 學年期資訊
                foreach (string moralScore in obj.TextScoreDic.Keys)
                {
                    Write(MoralScore2, moralScore);

                    //欄位名稱
                    foreach (string BehaviorConfigName1 in obj.TextScoreDic[moralScore].Keys)
                    {

                        //第二欄內容
                        if (BehaviorList2.Contains(BehaviorConfigName1))
                        {
                            Cell Score = GetMoveRightCell(MoralScore2, GetSummaryIndex(BehaviorConfigName1));
                            Write(Score, obj.TextScoreDic[moralScore][BehaviorConfigName1]);

                        }
                    }

                    Row Nextrow = MoralScore2.ParentRow.NextSibling as Row; //取得下一行
                    if (Nextrow == null)
                        break;
                    MoralScore2 = Nextrow.FirstCell; //第一格
                }
            }

            #endregion

            #region 日3

            if (BehaviorList3.Count > 0)
            {
                //填入日常生活表現3標題內容
                builder.MoveToMergeField("生1");
                setupCell = (Cell)builder.CurrentParagraph.ParentNode;

                foreach (string each in BehaviorList3)
                {
                    Write(setupCell, each); //填入導師評語

                    if (setupCell.NextSibling != null)
                    {
                        setupCell = setupCell.NextSibling as Cell; //取得下一格
                    }
                }

                Cell MoralScore3 = DicBack["日常生活表現欄3"];

                //KEY 學年期資訊
                foreach (string moralScore in obj.TextScoreDic.Keys)
                {
                    Write(MoralScore3, moralScore);

                    //欄位名稱
                    foreach (string BehaviorConfigName1 in obj.TextScoreDic[moralScore].Keys)
                    {
                        //第三欄內容
                        if (BehaviorList3.Contains(BehaviorConfigName1))
                        {
                            Cell Score = GetMoveRightCell(MoralScore3, GetSummaryIndex(BehaviorConfigName1));
                            Write(Score, obj.TextScoreDic[moralScore][BehaviorConfigName1]);
                        }
                    }

                    Row Nextrow = MoralScore3.ParentRow.NextSibling as Row; //取得下一行
                    if (Nextrow == null)
                        break;
                    MoralScore3 = Nextrow.FirstCell; //第一格
                }
            }

            #endregion

            #region 缺曠

            //填內容

            //builder.MoveToMergeField("統");
            Cell absenceCell2 = DicBack["缺曠資料欄1"];
            //RowNull2 = 0;
            //Cell MoralScore3 = (Cell)builder.CurrentParagraph.ParentNode;
            foreach (string absence3 in obj.DicAsbs.Keys)
            {
                //填入學期
                Write(absenceCell2, absence3);

                foreach (string SummaryName in obj.DicAsbs[absence3].Keys)
                {
                    if (!absenceList.Contains(SummaryName))
                        continue;

                    int index = GetSummaryIndex(SummaryName);

                    //如果是0,就是沒有值
                    if (index == 0)
                        continue;
                    //取得MoralScore3為基準的 index 格 cell
                    Cell MoralScore4 = GetMoveRightCell(absenceCell2, index);
                    //填入值
                    if (obj.DicAsbs[absence3][SummaryName] != "0")
                    {
                        Write(MoralScore4, obj.DicAsbs[absence3][SummaryName]);
                    }
                }


                Row Nextrow = absenceCell2.ParentRow.NextSibling as Row; //取得下一行
                if (Nextrow == null)
                    break;
                absenceCell2 = Nextrow.FirstCell; //第一格

            }
            #endregion

            #region 獎懲

            Cell MeritDemeritCell2 = DicBack["獎懲資料欄1"];
            foreach (string MeritFile in obj.DicMeritDemerit.Keys)
            {
                //填入學期
                Write(MeritDemeritCell2, MeritFile);

                foreach (string SummaryName in obj.DicMeritDemerit[MeritFile].Keys)
                {
                    int index = GetSummaryIndex(SummaryName);

                    //如果是0,就是沒有值
                    if (index == 0)
                        continue;

                    //取得MoralScore3為基準的 index 格 cell
                    Cell MoralScore4 = GetMoveRightCell(MeritDemeritCell2, index);

                    //填入值
                    if (obj.DicMeritDemerit[MeritFile][SummaryName] != "0")
                    {
                        Write(MoralScore4, obj.DicMeritDemerit[MeritFile][SummaryName]);
                    }
                }

                Row Nextrow = MeritDemeritCell2.ParentRow.NextSibling as Row; //取得下一行
                if (Nextrow == null)
                    break;
                MeritDemeritCell2 = Nextrow.FirstCell; //第一格
            }
            //builder.MoveToMergeField("獎");
            //Cell MeritDemeritCell = (Cell)builder.CurrentParagraph.ParentNode;

            //foreach (string moralScore in obj.SummaryDic.Keys)
            //{
            //    //填入學期
            //    Write(MeritDemeritCell, moralScore);

            //    //Cell
            //    int index = GetSummaryIndex("大功");
            //    Cell MeritDemerit = GetMoveRightCell(MeritDemeritCell, index);
            //    string MeritA = obj.SummaryDic[moralScore]["大功"];
            //    if (MeritA != "0")
            //        Write(MeritDemerit, MeritA);

            //    index = GetSummaryIndex("小功");
            //    MeritDemerit = GetMoveRightCell(MeritDemeritCell, index);
            //    string MeritB = obj.SummaryDic[moralScore]["小功"];
            //    if (MeritB != "0")
            //        Write(MeritDemerit, MeritB);

            //    index = GetSummaryIndex("嘉獎");
            //    MeritDemerit = GetMoveRightCell(MeritDemeritCell, index);
            //    string MeritC = obj.SummaryDic[moralScore]["嘉獎"];
            //    if (MeritC != "0")
            //        Write(MeritDemerit, MeritC);

            //    index = GetSummaryIndex("大過");
            //    MeritDemerit = GetMoveRightCell(MeritDemeritCell, index);
            //    string DemeritA = obj.SummaryDic[moralScore]["大過"];
            //    if (DemeritA != "0")
            //        Write(MeritDemerit, DemeritA);

            //    index = GetSummaryIndex("小過");
            //    MeritDemerit = GetMoveRightCell(MeritDemeritCell, index);
            //    string DemeritB = obj.SummaryDic[moralScore]["小過"];
            //    if (DemeritB != "0")
            //        Write(MeritDemerit, DemeritB);

            //    index = GetSummaryIndex("警告");
            //    MeritDemerit = GetMoveRightCell(MeritDemeritCell, index);
            //    string DemeritC = obj.SummaryDic[moralScore]["警告"];
            //    if (DemeritC != "0")
            //        Write(MeritDemerit, DemeritC);

            //    Row Nextrow = MeritDemeritCell.ParentRow.NextSibling as Row; //取得下一個Row
            //    if (Nextrow == null)
            //        break;
            //    MeritDemeritCell = Nextrow.FirstCell; //第一格Cell
            //    if (MeritDemeritCell == null)
            //        break;

            //}

            #endregion

            return PageOne;

            #endregion
        }

        /// <summary>
        /// 傳入Cell,定位名稱,數量,來建立Row的行數
        /// </summary>
        private void SetRowCount(string name, Cell cell, int p)
        {
            //記錄此Cell定位名稱
            DicBack.Add(name, cell);

            //取得目前Row
            Row 日3row = (Row)cell.ParentRow;

            //除了原來的Row-1,高於1就多建立幾行
            for (int x = 1; x < p; x++)
            {
                cell.ParentRow.ParentTable.InsertAfter(日3row.Clone(true), cell.ParentNode);
            }

        }

        /// <summary>
        /// 背景完成
        /// </summary>
        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MsgBox.Show("請設定節次資料!!");
                return;
            }

            if (e.Error != null)
            {
                MsgBox.Show("列印資料發生錯誤\n" + e.Error.Message);
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                return;
            }

            if (SingeFile)
            {
                #region 學生清單多檔列印
                //未完成 6/21
                //以學生資料列印報表名稱

                Dictionary<string, Document> DocDic = (Dictionary<string, Document>)e.Result;
                btnSave.Enabled = true;
                try
                {
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    fbd.Description = "請選擇學生日常表現總表檔案儲存位置\n規格為(學號_身分證號_班級_座號_姓名)";
                    fbd.ShowNewFolderButton = true;

                    if (fbd.ShowDialog() == DialogResult.Cancel)
                    {
                         MsgBox.Show("已取消存檔!!");
                         return;
                    }

                    //學號 報表名稱 學生姓名
                    //DocDic - 系統編號
                    //學號
                    //姓名等資料
                    //儲存路逕
                    Dictionary<string, Document> dic = new Dictionary<string, Document>();

                    foreach (string each in DocDic.Keys)
                    {
                        if (Data.DicStudent.ContainsKey(each))
                        {


                            StudentDataObj obj = Data.DicStudent[each];
                            Document inResult = DocDic[each];

                            StringBuilder sb = new StringBuilder();
                            sb.Append(fbd.SelectedPath + "\\");
                            sb.Append(obj.StudentRecord.StudentNumber + "_");
                            sb.Append(obj.StudentRecord.IDNumber + "_");
                            sb.Append((obj.StudentRecord.Class != null ? obj.StudentRecord.Class.Name : "") + "_");
                            sb.Append((obj.StudentRecord.SeatNo.HasValue ? "" + obj.StudentRecord.SeatNo.Value : "") + "_");
                            sb.Append(obj.StudentRecord.Name + ".doc");
                            if (!dic.ContainsKey(sb.ToString()))
                            {
                                dic.Add(sb.ToString(), inResult);
                            }
                            else
                            {
                                MsgBox.Show("發生檔案同名錯誤!!\n" + sb.ToString());
                            }
                        }
                    }

                    foreach (string each in dic.Keys)
                    {
                        dic[each].Save(each);
                    }
                    System.Diagnostics.Process.Start("explorer", fbd.SelectedPath);
                    MotherForm.SetStatusBarMessage("學生日常表現總表,儲存完成!!");
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤");
                    MotherForm.SetStatusBarMessage("檔案儲存錯誤");
                }
                #endregion
            }
            else
            {
                #region 學生清單單檔列印
                Document inResult = (Document)e.Result;
                btnSave.Enabled = true;

                try
                {
                    SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                    SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                    SaveFileDialog1.FileName = "學生日常表現總表";

                    if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        inResult.Save(SaveFileDialog1.FileName);
                        Process.Start(SaveFileDialog1.FileName);
                        MotherForm.SetStatusBarMessage("學生日常表現總表,列印完成!!");
                    }
                    else
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("已取消存檔");
                        return;
                    }
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                    MotherForm.SetStatusBarMessage("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                }
                #endregion
            }

        }

        /// <summary>
        /// 寫入資料
        /// </summary>
        private void Write(Cell cell, string text)
        {
            if (cell.FirstParagraph == null)
                cell.Paragraphs.Add(new Paragraph(cell.Document));
            cell.FirstParagraph.Runs.Clear();
            _run.Text = text;
            _run.Font.Size = 10;
            _run.Font.Name = "標楷體";
            cell.FirstParagraph.Runs.Add(_run.Clone(true));
        }

        /// <summary>
        /// 以Cell為基準,向右移一格
        /// </summary>
        private Cell GetMoveRightCell(Cell cell, int count)
        {
            if (count == 0) return cell;

            Row row = cell.ParentRow;
            int col_index = row.IndexOf(cell);
            Table table = row.ParentTable;
            int row_index = table.Rows.IndexOf(row);

            try
            {
                return table.Rows[row_index].Cells[col_index + count];
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// 日常生活表現設定值
        /// </summary>
        private List<string> GetBehaviorConfig()
        {
            List<string> TextScoreList = new List<string>();
            TextScoreList.Add("導師評語");
            foreach (SHSchool.Data.MoralityRecord each in SHSchool.Data.Morality.SelectAll())
            {
                if (!TextScoreList.Contains(each.Name))
                {
                    TextScoreList.Add(each.Name);
                }
            }

            //SmartSchool.Customization.Data.SystemInformation.getField("文字評量對照表");
            //System.Xml.XmlElement ElmTextScoreList = (System.Xml.XmlElement)SmartSchool.Customization.Data.SystemInformation.Fields["文字評量對照表"];
            //foreach (System.Xml.XmlNode Node in ElmTextScoreList.SelectNodes("Content/Morality"))
            //    TextScoreList.Add(Node.Attributes["Face"].InnerText);

            return TextScoreList;
        }

        /// <summary>
        /// 建立已知資料
        /// </summary>
        private void SetNameIndex()
        {
            DicSummaryIndex.Clear();
            DicSummaryIndex.Add("大功", 1);
            DicSummaryIndex.Add("小功", 2);
            DicSummaryIndex.Add("嘉獎", 3);
            DicSummaryIndex.Add("大過", 4);
            DicSummaryIndex.Add("小過", 5);
            DicSummaryIndex.Add("警告", 6);
            DicSummaryIndex.Add("留察", 7);

            UpdateCoddic.Clear();
            UpdateCoddic.Add("1", "新生");
            UpdateCoddic.Add("2", "畢業");
            UpdateCoddic.Add("3", "轉入");
            UpdateCoddic.Add("4", "轉出");
            UpdateCoddic.Add("5", "休學");
            UpdateCoddic.Add("6", "復學");
            UpdateCoddic.Add("7", "中輟");
            UpdateCoddic.Add("8", "續讀");
            UpdateCoddic.Add("9", "更正學籍");
        }

        /// <summary>
        /// 取得定義的統計資料Index
        /// </summary>
        private int GetSummaryIndex(string AttName)
        {
            if (DicSummaryIndex.ContainsKey(AttName))
            {
                return DicSummaryIndex[AttName];
            }
            else
            {
                return 0;
            }

        }

        /// <summary>
        /// 傳入異動代碼,取得異動原因
        /// </summary>
        private string GetUpdateRecordCode(string UpdateCode)
        {
            if (UpdateCoddic.ContainsKey(UpdateCode))
            {
                return UpdateCoddic[UpdateCode];
            }
            else
            {
                return "";
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //23節次以下
            SelectTypeForm typeForm = new SelectTypeForm(ConfigName);
            typeForm.ShowDialog();
        }
    }
}
