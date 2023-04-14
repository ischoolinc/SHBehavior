using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using Aspose.Cells;
using System.IO;
using System.Windows.Forms;
using SmartSchool.Customization.Data;
using System.Drawing;
using SmartSchool.Customization.Data.StudentExtension;
using SmartSchool.Common;
using FISCA.Presentation;

namespace SHSchool.Behavior
{
    class OverTheYearsStatistics
    {
        private BackgroundWorker _BGWTotalDisciplineAndAbsence;
        private bool _print_cleared;
        private Dictionary<string, List<string>> _print_types;
        private Dictionary<string, int> rowIndexTable = new Dictionary<string, int>();
        private int totalRow = 50;
        private int detailRow = 20;

        private Dictionary<string, string> PrintDic = new Dictionary<string, string>();

        public OverTheYearsStatistics(bool cleared, Dictionary<string, List<string>> types)
        {
            _print_cleared = cleared;
            _print_types = types;

            Check();

            _BGWTotalDisciplineAndAbsence = new BackgroundWorker();
            _BGWTotalDisciplineAndAbsence.DoWork += new DoWorkEventHandler(_BGWTotalDisciplineAndAbsence_DoWork);
            _BGWTotalDisciplineAndAbsence.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWTotalDisciplineAndAbsence_RunWorkerCompleted);
            _BGWTotalDisciplineAndAbsence.ProgressChanged += new ProgressChangedEventHandler(_BGWTotalDisciplineAndAbsence_ProgressChanged);
            _BGWTotalDisciplineAndAbsence.WorkerReportsProgress = true;
            _BGWTotalDisciplineAndAbsence.RunWorkerAsync();
        }

        private void Check()
        {
            foreach (SHSchool.Data.SHPeriodMappingInfo info in SHSchool.Data.SHPeriodMapping.SelectAll())
            {
                if (!PrintDic.ContainsKey(info.Name))
                {
                    PrintDic.Add(info.Name, info.Type);
                }
            }
        }

        void _BGWTotalDisciplineAndAbsence_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("歷年功過及出席統計產生中...", e.ProgressPercentage);
        }

        void _BGWTotalDisciplineAndAbsence_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            string reportName;
            string path;
            Workbook wb;

            object[] result = (object[])e.Result;
            reportName = (string)result[0];
            path = (string)result[1];
            wb = (Workbook)result[2];

            if (File.Exists(path))
            {
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
            }

            try
            {
                wb.Save(path);
                MotherForm.SetStatusBarMessage(reportName + "產生完成");
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".xlsx";
                sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        wb.Save(sd.FileName);
                    }
                    catch
                    {
                        MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        void _BGWTotalDisciplineAndAbsence_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "歷年功過及出席統計";

            _BGWTotalDisciplineAndAbsence.ReportProgress(0);

            #region 取得資料

            AccessHelper helper = new AccessHelper();
            List<StudentRecord> students = helper.StudentHelper.GetSelectedStudent();

            helper.StudentHelper.FillSemesterEntryScore(true, students);
            helper.StudentHelper.FillReward(students);
            helper.StudentHelper.FillAttendance(students);

            StatisticsCollection stat = new StatisticsCollection();
            Dictionary<string, List<string>> detailList = new Dictionary<string, List<string>>();

            Dictionary<string, int> rewardDict = new Dictionary<string, int>();

            foreach (StudentRecord each in students)
            {
                StatisticsData data = new StatisticsData();
                List<string> list = new List<string>();

                //獎懲
                foreach (RewardInfo info in each.RewardList)
                {
                    int schoolyear = info.SchoolYear;
                    int semester = info.Semester;

                    if (info.UltimateAdmonition == true)
                    {
                        data.AddItem(schoolyear, semester, "留校察看", 1);
                        list.Add(CDATE(info.OccurDate.ToShortDateString()) + " " + info.OccurReason + " 留校察看");
                        continue;
                    }

                    rewardDict.Clear();

                    data.AddItem(schoolyear, semester, "大功", (decimal)info.AwardA);
                    data.AddItem(schoolyear, semester, "小功", (decimal)info.AwardB);
                    data.AddItem(schoolyear, semester, "嘉獎", (decimal)info.AwardC);
                    rewardDict.Add("大功", info.AwardA);
                    rewardDict.Add("小功", info.AwardB);
                    rewardDict.Add("嘉獎", info.AwardC);

                    if (_print_cleared == true || info.Cleared == false)
                    {
                        data.AddItem(schoolyear, semester, "大過", (decimal)info.FaultA);
                        data.AddItem(schoolyear, semester, "小過", (decimal)info.FaultB);
                        data.AddItem(schoolyear, semester, "警告", (decimal)info.FaultC);
                        rewardDict.Add("大過", info.FaultA);
                        rewardDict.Add("小過", info.FaultB);
                        rewardDict.Add("警告", info.FaultC);
                    }

                    string rewardStat = "";
                    foreach (string var in new string[] { "大功", "小功", "嘉獎", "大過", "小過", "警告" })
                    {
                        if (rewardDict.ContainsKey(var) && rewardDict[var] > 0)
                        {
                            if (!string.IsNullOrEmpty(rewardStat))
                                rewardStat += ", ";
                            rewardStat += var + rewardDict[var] + "次";
                        }
                    }

                    //兩個條件可加入明細清單
                    //1.如果勾選消過也列印
                    //2.資料未消過
                    if (_print_cleared == true)
                    {
                        string all = "";
                        if (info.Cleared)
                        {
                            all = CDATE(info.OccurDate.ToShortDateString()) + " " + info.OccurReason + " " + rewardStat + " (已銷過)";
                        }
                        else
                        {
                            all = CDATE(info.OccurDate.ToShortDateString()) + " " + info.OccurReason + " " + rewardStat;
                        }
                        list.Add(all);
                    }
                    else if (info.Cleared == false)
                    {
                        string all = CDATE(info.OccurDate.ToShortDateString()) + " " + info.OccurReason + " " + rewardStat;
                        list.Add(all);
                    }

                    #region 註解掉了
                    //if (((info.OccurReason + " " + rewardStat) as string).Length >= 20)
                    //{
                    //    List<string> mini = new List<string>();
                    //    string reason = info.OccurReason;

                    //    while (reason.Length >= 20)
                    //    {
                    //        mini.Add(reason.Substring(0, 19));
                    //        reason = reason.Substring(19);
                    //    }
                    //    if (((reason + " " + rewardStat) as string).Length >= 20)
                    //    {
                    //        mini.Add(reason);
                    //        mini.Add(rewardStat);
                    //    }
                    //    else
                    //        mini.Add(reason + " " + rewardStat);

                    //    //mini.Add(CDATE(info.OccurDate.ToShortDateString()) + " ");

                    //    for (int i = 0; i < mini.Count; i++)
                    //    {
                    //        if (i == 0)
                    //            list.Add(CDATE(info.OccurDate.ToShortDateString()) + " " + mini[i]);
                    //        else
                    //            list.Add("　　　" + mini[i]);
                    //    }
                    //}
                    //else
                    //    list.Add(all);
                    #endregion
                }
                detailList.Add(each.StudentID, list);

                //缺曠
                foreach (AttendanceInfo info in each.AttendanceList)
                {
                    int schoolyear = info.SchoolYear;
                    int semester = info.Semester;
                    if (PrintDic.ContainsKey(info.Period))
                    {
                        data.AddItem(schoolyear, semester, PrintDic[info.Period] + "_" + info.Absence, 1);
                    }
                }

                //德行成績
                foreach (SemesterEntryScoreInfo info in each.SemesterEntryScoreList)
                {
                    if (info.Entry == "德行")
                    {
                        int schoolyear = info.SchoolYear;
                        int semester = info.Semester;

                        data.AddItem(schoolyear, semester, "德行成績", info.Score);
                    }
                }

                data.MendSemester();

                stat.Add(each.StudentID, data);
            }

            #endregion

            #region 產生範本

            int AsbCount = 0;
            foreach (string each in _print_types.Keys)
            {
                AsbCount += _print_types[each].Count;
            }

            totalRow = detailRow + 11 + AsbCount;

            foreach (StudentRecord each in students)
            {
                if (detailRow < detailList[each.StudentID].Count / 2 + 1)
                {
                    detailRow = detailList[each.StudentID].Count / 2 + 1;

                    totalRow = detailRow + 11 + AsbCount;
                }
            }


            Workbook pt1 = GetTemplate(6);
            Workbook pt2 = GetTemplate(8);
            Workbook pt3 = GetTemplate(10);

            Dictionary<int, Workbook> pts = new Dictionary<int, Workbook>();
            pts.Add(6, pt1);
            pts.Add(8, pt2);
            pts.Add(10, pt3);

            Range ptr1 = pt1.Worksheets[0].Cells.CreateRange(0, totalRow, false);
            Range ptr2 = pt2.Worksheets[0].Cells.CreateRange(0, totalRow, false);
            Range ptr3 = pt3.Worksheets[0].Cells.CreateRange(0, totalRow, false);

            Dictionary<int, Range> ptrs = new Dictionary<int, Range>();
            ptrs.Add(6, ptr1);
            ptrs.Add(8, ptr2);
            ptrs.Add(10, ptr3);

            #endregion

            #region 產生報表

            Workbook wb = new Workbook();
            wb.Copy(pt1);
            wb.CopyTheme(pt1);

            wb.Worksheets[0].Copy(pt1.Worksheets[0]);

            int wsCount = wb.Worksheets.Count;
            for (int i = wsCount - 1; i > 0; i--)
                wb.Worksheets.RemoveAt(i);
            wb.Worksheets[0].Name = "一般(6學期)";

            Dictionary<int, int> sheetIndex = new Dictionary<int, int>();
            Dictionary<int, Worksheet> sheets = new Dictionary<int, Worksheet>();
            sheets.Add(6, wb.Worksheets[0]);
            sheetIndex.Add(6, 0);

            Worksheet ws = wb.Worksheets[0];

            int index = 0;
            int cur = 6;

            int pages = 500;

            //
            int start = 1;
            int limit = pages;

            int studentCount = 0;

            //學生數
            int allStudentCount = students.Count;

            foreach (StudentRecord each in students)
            {
                //取得學生統計
                StatisticsData data = stat[each.StudentID];

                #region 判斷學期數
                if (data.GetSemesterNumber() > 6 && data.GetSemesterNumber() <= 8)
                {
                    //學期數大於6小於8
                    //特殊報表
                    if (!sheets.ContainsKey(8))
                    {
                        int new_ws_index = wb.Worksheets.Add();
                        wb.Worksheets[new_ws_index].Copy(pt2.Worksheets[0]);

                        wb.Worksheets[new_ws_index].Name = "特殊(8學期)";
                        sheets.Add(8, wb.Worksheets[new_ws_index]);
                        sheetIndex.Add(8, 0);
                    }

                    //變更預設報表位置
                    ws = sheets[8];
                    //取得報表定位
                    index = sheetIndex[8];
                    //取得報表定位?
                    cur = 8;
                    studentCount++;
                }
                else if (data.GetSemesterNumber() == 6)
                {

                    ws = sheets[6];
                    index = sheetIndex[6];
                    cur = 6;
                    studentCount++;
                }
                else
                {
                    //學期數大於8小於10
                    //特殊1報表
                    if (!sheets.ContainsKey(10))
                    {
                        int new_ws_index = wb.Worksheets.Add();
                        wb.Worksheets[new_ws_index].Copy(pt3.Worksheets[0]);
                        wb.Worksheets[new_ws_index].Name = "特殊(其它)";
                        sheets.Add(10, wb.Worksheets[new_ws_index]);
                        sheetIndex.Add(10, 0);
                    }
                    ws = sheets[10];
                    index = sheetIndex[10];
                    cur = 10;
                    studentCount++;
                }
                #endregion

                //取得畫面範圍 Range(0~TotalRow)
                ws.Cells.CreateRange(index, totalRow, false).Copy(ptrs[cur]);
                ws.Cells.CreateRange(index, totalRow, false).CopyData(ptrs[cur]);
                ws.Cells.CreateRange(index, totalRow, false).CopyStyle(ptrs[cur]);

                //填入標題
                ws.Cells[index + 1, 0].PutValue(string.Format("班級：{0}　　　　學號：{1}　　　　姓名：{2}", (each.RefClass != null) ? each.RefClass.ClassName : "", each.StudentNumber, each.StudentName));

                //第一行位置(標題+2)
                int firstRow = index + 2;

                //第一列位置
                int col = 2;

                //處理每個學期的統計
                foreach (string semsString in data.Semesters)
                {

                    //填入學年度/學期
                    ws.Cells[firstRow, col].PutValue(DisplaySemester(semsString));

                    //取得學期統計資料
                    Dictionary<string, decimal> semsDict = data.GetItem(semsString);

                    foreach (string item in semsDict.Keys)
                    {
                        if (rowIndexTable.ContainsKey(item))
                            ws.Cells[index + rowIndexTable[item], col].PutValue((semsDict[item] <= 0) ? "" : semsDict[item].ToString());
                    }

                    col++;
                }

                if (rowIndexTable.ContainsKey("明細"))
                {
                    //資料一律由第二列進行列印
                    int detailColIndex = 2;

                    //
                    int detailIndex = index + rowIndexTable["明細"];

                    //明細定位點
                    int detailCount = 0;

                    //取得學生明細清單
                    foreach (string var in detailList[each.StudentID])
                    {
                        //明細定位點+1
                        detailCount++;
                        //定位點大於明細畫面範圍時
                        //將換為第二區進行列印
                        if (detailCount > detailRow)
                        {
                            detailColIndex += (cur / 2);
                            detailIndex = index + rowIndexTable["明細"];
                            detailCount = 0;
                        }
                        ws.Cells[detailIndex++, detailColIndex].PutValue(var);
                    }
                }

                //?
                index += totalRow;
                sheetIndex[cur] = index;

                //增加列印分割線
                ws.HorizontalPageBreaks.Add(index, 0);

                //當學生大於極限值 , 並且列印完所有學生
                if (studentCount >= limit && studentCount < allStudentCount)
                {
                    //取得工作表名稱
                    //string orig_name = ws.Name;

                    //
                    //ws.Name = ws.Name + " (" + start + " ~ " + studentCount + ")";
                    ws = wb.Worksheets[wb.Worksheets.Add()];
                    ws.Copy(pts[cur].Worksheets[0]);

                    //ws.Name = orig_name;
                    start += pages;
                    limit += pages;
                    index = 0;
                    sheetIndex[cur] = index;
                    sheets[cur] = ws;
                }

                //回報進度
                _BGWTotalDisciplineAndAbsence.ReportProgress((int)(((double)studentCount * 100.0) / (double)allStudentCount));
            }

            //if (cur == 6)
            //ws.Name = ws.Name + " (" + start + " ~ " + studentCount + ")";

            #endregion

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xlsx");
            e.Result = new object[] { reportName, path, wb };
        }

        private string CDATE(string p)
        {
            DateTime d = DateTime.Now;
            if (p != "" && DateTime.TryParse(p, out d))
            {
                return "" + (d.Year - 1911) + "/" + d.Month + "/" + d.Day;
            }
            else
                return "";
        }

        private string DisplaySemester(string sems)
        {
            string[] split = sems.Split('_');
            string schoolyear = split[0];
            string semester = split[1];
            string semesterDisplay = "";
            if (semester == "1")
                semesterDisplay = "上";
            else if (semester == "2")
                semesterDisplay = "下";

            return schoolyear + semesterDisplay;
        }

        private Workbook GetTemplate(int sems_num)
        {
            Workbook template = new Workbook();
            template = new Workbook(new MemoryStream(Properties.Resources.歷年功過及出席統計), new LoadOptions(LoadFormat.Excel97To2003));

            Worksheet tempWorksheet = template.Worksheets["" + sems_num];

            Range tempHeader = tempWorksheet.Cells.CreateRange(0, 2, false);
            Range tempRow = tempWorksheet.Cells.CreateRange(2, 1, false);
            Range tempStatistics = tempWorksheet.Cells.CreateRange(3, 1, false);
            Range tempDetail = tempWorksheet.Cells.CreateRange(5, 1, false);

            Workbook prototype = new Workbook();
            prototype.Copy(template);
            prototype.CopyTheme(template);

            int count = prototype.Worksheets.Count;
            for (int i = count - 1; i > 0; i--)
                prototype.Worksheets.RemoveAt(i);

            prototype.Worksheets[0].Copy(tempWorksheet);
            prototype.Worksheets[0].Cells.CreateRange(0, 2, false).Copy(tempHeader);
            prototype.Worksheets[0].Cells.CreateRange(0, 2, false).CopyData(tempHeader);
            prototype.Worksheets[0].Cells.CreateRange(0, 2, false).CopyStyle(tempHeader);

            int index = 2;
            rowIndexTable.Clear();

            //標題
            prototype.Worksheets[0].Cells[0, 0].PutValue(SmartSchool.Customization.Data.SystemInformation.SchoolChineseName + "學生個人歷年功過及出席統計表");

            //學期
            prototype.Worksheets[0].Cells.CreateRange(index, 1, false).Copy(tempRow);
            prototype.Worksheets[0].Cells.CreateRange(index, 1, false).CopyData(tempRow);
            prototype.Worksheets[0].Cells.CreateRange(index, 1, false).CopyStyle(tempRow);

            prototype.Worksheets[0].Cells[index, 0].PutValue("學年度學期");
            index++;

            //獎懲部分
            foreach (string var in new string[] { "大功", "小功", "嘉獎" })
            {
                prototype.Worksheets[0].Cells.CreateRange(index, 1, false).Copy(tempStatistics);
                prototype.Worksheets[0].Cells.CreateRange(index, 1, false).CopyData(tempStatistics);
                prototype.Worksheets[0].Cells.CreateRange(index, 1, false).CopyStyle(tempStatistics);

                prototype.Worksheets[0].Cells[index, 1].PutValue(var);

                rowIndexTable.Add(var, index);
                index++;
            }
            foreach (string var in new string[] { "大過", "小過", "警告", "留校察看" })
            {
                prototype.Worksheets[0].Cells.CreateRange(index, 1, false).Copy(tempStatistics);
                prototype.Worksheets[0].Cells.CreateRange(index, 1, false).CopyData(tempStatistics);
                prototype.Worksheets[0].Cells.CreateRange(index, 1, false).CopyStyle(tempStatistics);

                prototype.Worksheets[0].Cells[index, 1].PutValue(var);

                rowIndexTable.Add(var, index);
                index++;
            }

            prototype.Worksheets[0].Cells.CreateRange(index - 7, 0, 3, 1).Merge();
            prototype.Worksheets[0].Cells[index - 7, 0].PutValue("獎勵");
            prototype.Worksheets[0].Cells.CreateRange(index - 4, 0, 4, 1).Merge();
            prototype.Worksheets[0].Cells[index - 4, 0].PutValue("懲戒");

            //缺曠部分
            foreach (string period in _print_types.Keys)
            {
                foreach (string absence in _print_types[period])
                {
                    prototype.Worksheets[0].Cells.CreateRange(index, 1, false).Copy(tempStatistics);
                    prototype.Worksheets[0].Cells.CreateRange(index, 1, false).CopyData(tempStatistics);
                    prototype.Worksheets[0].Cells.CreateRange(index, 1, false).CopyStyle(tempStatistics);

                    prototype.Worksheets[0].Cells[index, 1].PutValue(absence);

                    rowIndexTable.Add(period + "_" + absence, index);
                    index++;
                }
                prototype.Worksheets[0].Cells.CreateRange(index - _print_types[period].Count, 0, _print_types[period].Count, 1).Merge();
                prototype.Worksheets[0].Cells[index - _print_types[period].Count, 0].PutValue(period);
            }

            //德行成績部分
            prototype.Worksheets[0].Cells.CreateRange(index, 1, false).Copy(tempRow);
            prototype.Worksheets[0].Cells.CreateRange(index, 1, false).CopyData(tempRow);
            prototype.Worksheets[0].Cells.CreateRange(index, 1, false).CopyStyle(tempRow);

            prototype.Worksheets[0].Cells[index, 0].PutValue("德行成績");

            rowIndexTable.Add("德行成績", index);
            index++;


            //獎懲明細部分

            ////檢查每個學生獎懲明細是否會超過表格大小
            //foreach (string var in disciplineDetail.Keys)
            //{
            //    if (disciplineDetail[var].Count > (totalRows - index) * 2)
            //    {
            //        totalRows += (disciplineDetail[var].Count - (totalRows - index) * 2) / 2;
            //        totalRows += (disciplineDetail[var].Count - (totalRows - index) * 2) % 2;
            //    }
            //}

            rowIndexTable.Add("明細", index);

            //if ((index + detailRow) > totalRow)
            //{
            //    detailRow = totalRow - index;
            //}

            for (int j = index; j < index + detailRow; j++)
            {
                prototype.Worksheets[0].Cells.CreateRange(j, 1, false).Copy(tempDetail);
                prototype.Worksheets[0].Cells.CreateRange(j, 1, false).CopyStyle(tempDetail);
                prototype.Worksheets[0].Cells.CreateRange(j, 1, false).CopyData(tempDetail);
            }
            prototype.Worksheets[0].Cells.CreateRange(index, 0, detailRow, 2).Merge();

            prototype.Worksheets[0].Cells[index, 0].PutValue("獎懲明細");

            //加上報表底線
            prototype.Worksheets[0].Cells.CreateRange(index + detailRow - 1, 0, 1, sems_num + 2).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

            return prototype;
        }

        class StatisticsCollection : Dictionary<string, StatisticsData>
        {
            #region 註解掉
            //public List<int> CheckSemesterNumber()
            //{
            //    List<int> sems = new List<int>();
            //    foreach (StatisticsData data in this.Values)
            //    {
            //        int count = data.GetSemesterNumber();
            //        if (!sems.Contains(count))
            //            sems.Add(count);
            //    }
            //    sems.Sort();
            //    return sems;
            //}
            #endregion
        }

        class StatisticsData
        {
            private Dictionary<string, Dictionary<string, decimal>> _data = new Dictionary<string, Dictionary<string, decimal>>();
            public Dictionary<string, Dictionary<string, decimal>> Data
            {
                get { return _data; }
            }

            public List<string> Semesters
            {
                get
                {
                    List<string> sorted = new List<string>();
                    sorted.AddRange(_data.Keys);
                    sorted.Sort(SortBySchool);

                    return sorted;
                }
            }

            private int SortBySchool(string g, string k)
            {
                return g.PadLeft(10, '0').CompareTo(k.PadLeft(10, '0'));
            }

            public void AddItem(int schoolyear, int semester, string data_key, decimal data_value)
            {
                string ss = schoolyear + "_" + semester;
                if (!_data.ContainsKey(ss))
                    _data.Add(ss, new Dictionary<string, decimal>());
                if (!_data[ss].ContainsKey(data_key))
                    _data[ss].Add(data_key, data_value);
                else
                    _data[ss][data_key] += data_value;
            }

            public Dictionary<string, decimal> GetItem(string ss)
            {
                if (_data.ContainsKey(ss))
                    return _data[ss];
                return new Dictionary<string, decimal>();
            }

            public void MendSemester()
            {
                List<string> mend = new List<string>();

                foreach (string ss in _data.Keys)
                {
                    string[] split = ss.Split('_');
                    string schoolyear = split[0];
                    string semester = split[1];

                    if (semester == "1")
                    {
                        string new_key = schoolyear + "_2";
                        if (!_data.ContainsKey(new_key) && !mend.Contains(new_key))
                            mend.Add(new_key);
                    }
                    else if (semester == "2")
                    {
                        string new_key = schoolyear + "_1";
                        if (!_data.ContainsKey(new_key) && !mend.Contains(new_key))
                            mend.Add(new_key);
                    }
                }

                foreach (string key in mend)
                {
                    if (!_data.ContainsKey(key))
                        _data.Add(key, new Dictionary<string, decimal>());
                }
            }

            public int GetSemesterNumber()
            {
                return _data.Keys.Count;
            }
        }
    }
}
