using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Aspose.Cells;
using FISCA.DSAUtil;
using Framework;
using K12.Data;
using SHSchool.Data;

namespace SHSchool.Behavior.StudentExtendControls.Reports.學生獎勵明細
{
    internal class Report : IReport
    {
        #region IReport 成員

        private BackgroundWorker _BGWDisciplineDetail;
        SelectMeritForm form;

        public void Print()
        {

            if (K12.Presentation.NLDPanels.Student.SelectedSource.Count == 0)
                return;

            //警告使用者別做傻事
            if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 1500)
            {
                MsgBox.Show("您選取的學生超過 1500 個，可能會發生意想不到的錯誤，請減少選取的學生。");
                return;
            }

            form = new SelectMeritForm();

            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化學生獎勵記錄明細...");

                //object[] args = new object[] { form.SchoolYear, form.Semester };

                _BGWDisciplineDetail = new BackgroundWorker();
                _BGWDisciplineDetail.DoWork += new DoWorkEventHandler(_BGWDisciplineDetail_DoWork);
                _BGWDisciplineDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.ExcelReport_RunWorkerCompleted);
                _BGWDisciplineDetail.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWDisciplineDetail.WorkerReportsProgress = true;
                _BGWDisciplineDetail.RunWorkerAsync();
            }
        }

        #endregion

        void _BGWDisciplineDetail_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "學生獎勵明細";

            #region 快取相關資料

            //選擇的學生
            List<SHStudentRecord> selectedStudents = SHStudent.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
            selectedStudents.Sort(new Comparison<SHStudentRecord>(CommonMethods.SHClassSeatNoComparer));

            //紀錄所有學生ID
            List<string> allStudentID = new List<string>();

            //每一位學生的獎勵明細
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> studentDisciplineDetail = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            //每一位學生的獎勵累計資料
            Dictionary<string, Dictionary<string, int>> studentDisciplineStatistics = new Dictionary<string, Dictionary<string, int>>();

            //紀錄每一種獎勵在報表中的 column index
            Dictionary<string, int> columnTable = new Dictionary<string, int>();

            //取得所有學生ID
            foreach (SHStudentRecord var in selectedStudents)
            {
                allStudentID.Add(var.ID);
            }

            //對照表
            Dictionary<string, string> meritTable = new Dictionary<string, string>();
            meritTable.Add("A", "大功");
            meritTable.Add("B", "小功");
            meritTable.Add("C", "嘉獎");

            //初始化
            string[] columnString = new string[] { "嘉獎", "小功", "大功", "事由", "備註" };
            int i = 4;
            foreach (string s in columnString)
            {
                columnTable.Add(s, i++);
            }

            //產生 DSRequest，取得缺曠明細

            DSResponse dsrsp;

            if (form.SelectDayOrSchoolYear) //依日期
            {
                if (form.SetupTime) //依發生日期
                {
                    #region 依發生日期
                    DSXmlHelper helper = new DSXmlHelper("Request");
                    helper.AddElement("Field");
                    helper.AddElement("Field", "All");
                    helper.AddElement("Condition");
                    foreach (string var in allStudentID)
                    {
                        helper.AddElement("Condition", "RefStudentID", var);
                    }

                    helper.AddElement("Condition", "StartDate", form.StartDay);
                    helper.AddElement("Condition", "EndDate", form.EndDay);

                    helper.AddElement("Order");
                    helper.AddElement("Order", "OccurDate", "asc");
                    dsrsp = QueryDiscipline.GetDiscipline(new DSRequest(helper));
                    #endregion
                }
                else //依登錄日期
                {
                    #region 依登錄日期
                    DSXmlHelper helper = new DSXmlHelper("Request");
                    helper.AddElement("Field");
                    helper.AddElement("Field", "All");
                    helper.AddElement("Condition");
                    foreach (string var in allStudentID)
                    {
                        helper.AddElement("Condition", "RefStudentID", var);
                    }

                    helper.AddElement("Condition", "StartRegisterDate", form.StartDay);
                    helper.AddElement("Condition", "EndRegisterDate", form.EndDay);

                    helper.AddElement("Order");
                    helper.AddElement("Order", "OccurDate", "asc");
                    dsrsp = QueryDiscipline.GetDiscipline(new DSRequest(helper));
                    #endregion
                }
            }
            else //依學期
            {
                if (form.checkBoxX1Bool) //全部學期列印
                {
                    #region 全部學期列印
                    DSXmlHelper helper = new DSXmlHelper("Request");
                    helper.AddElement("Field");
                    helper.AddElement("Field", "All");
                    helper.AddElement("Condition");
                    foreach (string var in allStudentID)
                    {
                        helper.AddElement("Condition", "RefStudentID", var);
                    }
                    helper.AddElement("Order");
                    helper.AddElement("Order", "OccurDate", "asc");
                    dsrsp = QueryDiscipline.GetDiscipline(new DSRequest(helper));
                    #endregion
                }
                else //指定學期列印
                {
                    #region 指定學期列印
                    DSXmlHelper helper = new DSXmlHelper("Request");
                    helper.AddElement("Field");
                    helper.AddElement("Field", "All");
                    helper.AddElement("Condition");
                    foreach (string var in allStudentID)
                    {
                        helper.AddElement("Condition", "RefStudentID", var);
                    }

                    helper.AddElement("Condition", "SchoolYear", form.SchoolYear);
                    helper.AddElement("Condition", "Semester", form.Semester);

                    helper.AddElement("Order");
                    helper.AddElement("Order", "OccurDate", "asc");
                    dsrsp = QueryDiscipline.GetDiscipline(new DSRequest(helper));
                    #endregion
                }
            }

            if (dsrsp == null)
            {
                MsgBox.Show("未取得獎勵資料");
                return;
            }

            foreach (XmlElement var in dsrsp.GetContent().GetElements("Discipline"))
            {
                if (var.SelectSingleNode("MeritFlag").InnerText == "1")
                {
                    string studentID = var.SelectSingleNode("RefStudentID").InnerText;
                    string schoolYear = var.SelectSingleNode("SchoolYear").InnerText;
                    string semester = var.SelectSingleNode("Semester").InnerText;
                    string occurDate = DateTime.Parse(var.SelectSingleNode("OccurDate").InnerText).ToShortDateString();
                    string reason = var.SelectSingleNode("Reason").InnerText;
                    string remark = var.SelectSingleNode("Remark").InnerText; //2019/12/31 - 新增
                    string disciplineID = var.GetAttribute("ID");
                    string sso = schoolYear + "_" + semester + "_" + occurDate + "_" + disciplineID;

                    //初始化累計資料
                    if (!studentDisciplineStatistics.ContainsKey(studentID))
                        studentDisciplineStatistics.Add(studentID, new Dictionary<string, int>());

                    //每一位學生獎勵資料
                    if (!studentDisciplineDetail.ContainsKey(studentID))
                        studentDisciplineDetail.Add(studentID, new Dictionary<string, Dictionary<string, string>>());
                    if (!studentDisciplineDetail[studentID].ContainsKey(sso))
                        studentDisciplineDetail[studentID].Add(sso, new Dictionary<string, string>());

                    //加入事由
                    if (!studentDisciplineDetail[studentID][sso].ContainsKey("事由"))
                        studentDisciplineDetail[studentID][sso].Add("事由", reason);

                    //加入事由
                    if (!studentDisciplineDetail[studentID][sso].ContainsKey("備註"))
                        studentDisciplineDetail[studentID][sso].Add("備註", remark);

                    XmlElement discipline = (XmlElement)var.SelectSingleNode("Detail/Discipline/Merit");
                    foreach (XmlAttribute attr in discipline.Attributes)
                    {
                        if (meritTable.ContainsKey(attr.Name))
                        {
                            string name = meritTable[attr.Name];

                            if (!studentDisciplineStatistics[studentID].ContainsKey(name))
                                studentDisciplineStatistics[studentID].Add(name, 0);

                            int v;

                            if (int.TryParse(attr.InnerText, out v))
                                studentDisciplineStatistics[studentID][name] += v;

                            if (!studentDisciplineDetail[studentID][sso].ContainsKey(name))
                                studentDisciplineDetail[studentID][sso].Add(name, attr.InnerText);
                        }
                    }
                }
            }

            #endregion

            #region 產生範本

            Workbook template = new Workbook();
            template.Open(new MemoryStream(Properties.Resources.學生獎勵記錄明細), FileFormatType.Excel2003);

            Workbook prototype = new Workbook();
            prototype.Copy(template);

            Worksheet ptws = prototype.Worksheets[0];

            int startPage = 1;
            int pageNumber = 1;

            int columnNumber = 9;

            //合併標題列
            ptws.Cells.CreateRange(0, 0, 1, columnNumber).Merge();
            ptws.Cells.CreateRange(1, 0, 1, columnNumber).Merge();

            Range ptHeader = ptws.Cells.CreateRange(0, 4, false);
            Range ptEachRow = ptws.Cells.CreateRange(4, 1, false);

            #endregion

            #region 產生報表

            Workbook wb = new Workbook();
            wb.Copy(prototype);
            Worksheet ws = wb.Worksheets[0];

            int index = 0;
            int dataIndex = 0;

            int studentCount = 1;

            foreach (SHStudentRecord studentInfo in selectedStudents)
            {
                //回報進度
                _BGWDisciplineDetail.ReportProgress((int)(((double)studentCount++ * 100.0) / (double)selectedStudents.Count));

                if (!studentDisciplineDetail.ContainsKey(studentInfo.ID))
                    continue;

                //如果不是第一頁，就在上一頁的資料列下邊加黑線
                if (index != 0)
                    ws.Cells.CreateRange(index - 1, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

                //複製 Header
                ws.Cells.CreateRange(index, 4, false).Copy(ptHeader);

                //填寫基本資料
                ws.Cells[index, 0].PutValue(School.ChineseName + "\n個人獎勵明細");
                ws.Cells[index + 1, 0].PutValue("班級：" + ((studentInfo.Class == null) ? "　　　" : studentInfo.Class.Name) + "　　座號：" + ((studentInfo.SeatNo == null) ? "　" : studentInfo.SeatNo.ToString()) + "　　姓名：" + studentInfo.Name + "　　學號：" + studentInfo.StudentNumber);

                dataIndex = index + 4;
                int recordCount = 0;

                Dictionary<string, Dictionary<string, string>> disciplineDetail = studentDisciplineDetail[studentInfo.ID];

                foreach (string sso in disciplineDetail.Keys)
                {
                    string[] ssoSplit = sso.Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);

                    //複製每一個 row
                    ws.Cells.CreateRange(dataIndex, 1, false).Copy(ptEachRow);

                    //填寫學生獎勵資料
                    ws.Cells[dataIndex, 0].PutValue(ssoSplit[0]);
                    ws.Cells[dataIndex, 1].PutValue(ssoSplit[1]);
                    ws.Cells[dataIndex, 2].PutValue(ssoSplit[2]);
                    ws.Cells[dataIndex, 3].PutValue(CommonMethods.GetChineseDayOfWeek(DateTime.Parse(ssoSplit[2])));

                    Dictionary<string, string> record = disciplineDetail[sso];
                    foreach (string name in record.Keys)
                    {
                        if (meritTable.ContainsValue(name))
                        {
                            int v;

                            if (int.TryParse(record[name], out v))
                            {
                                if (v > 0)
                                    ws.Cells[dataIndex, columnTable[name]].PutValue(record[name]);
                            }
                        }
                        else
                        {
                            if (columnTable.ContainsKey(name))
                                ws.Cells[dataIndex, columnTable[name]].PutValue(record[name]);
                        }
                    }

                    dataIndex++;
                    recordCount++;
                }

                //獎懲統計資訊
                Range disciplineStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber);
                disciplineStatisticsRange.Copy(ptEachRow);
                disciplineStatisticsRange.Merge();
                disciplineStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Double, Color.Black);
                disciplineStatisticsRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
                disciplineStatisticsRange.RowHeight = 20.0;
                ws.Cells[dataIndex, 0].Style.HorizontalAlignment = TextAlignmentType.Center;
                ws.Cells[dataIndex, 0].Style.VerticalAlignment = TextAlignmentType.Center;
                ws.Cells[dataIndex, 0].Style.Font.Size = 10;
                ws.Cells[dataIndex, 0].PutValue("獎勵總計");
                dataIndex++;

                //獎懲統計內容
                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).Copy(ptEachRow);
                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).RowHeight = 27.0;
                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).Merge();
                ws.Cells[dataIndex, 0].Style.HorizontalAlignment = TextAlignmentType.Center;
                ws.Cells[dataIndex, 0].Style.VerticalAlignment = TextAlignmentType.Center;
                ws.Cells[dataIndex, 0].Style.Font.Size = 10;
                ws.Cells[dataIndex, 0].Style.ShrinkToFit = true;

                StringBuilder text = new StringBuilder("");
                Dictionary<string, int> disciplineStatistics = studentDisciplineStatistics[studentInfo.ID];

                foreach (string type in disciplineStatistics.Keys)
                {
                    if (disciplineStatistics[type] > 0)
                    {
                        if (text.ToString() != "")
                            text.Append("　");
                        text.Append(type + "：" + disciplineStatistics[type]);
                    }
                }

                ws.Cells[dataIndex, 0].PutValue(text.ToString());

                dataIndex++;

                //資料列上邊加上黑線
                ws.Cells.CreateRange(index + 3, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

                //表格最右邊加上黑線
                ws.Cells.CreateRange(index + 2, columnNumber - 1, recordCount + 4, 1).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);

                index = dataIndex;

                //設定分頁
                if (pageNumber < 500)
                {
                    ws.HPageBreaks.Add(index, columnNumber);
                    pageNumber++;
                }
                else
                {
                    ws.Name = startPage + " ~ " + (pageNumber + startPage - 1);
                    ws = wb.Worksheets[wb.Worksheets.Add()];
                    ws.Copy(prototype.Worksheets[0]);
                    startPage += pageNumber;
                    pageNumber = 1;
                    index = 0;
                }

            }


            if (dataIndex > 0)
            {
                //最後一頁的資料列下邊加上黑線
                ws.Cells.CreateRange(dataIndex - 1, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
                ws.Name = startPage + " ~ " + (pageNumber + startPage - 2);
            }
            else
                wb = new Workbook();


            #endregion

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xlt");
            e.Result = new object[] { reportName, path, wb };
        }
    }
}
