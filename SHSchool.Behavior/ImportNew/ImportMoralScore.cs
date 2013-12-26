using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.Import;
using System.Xml;
using K12.Data;
using FISCA.DSAUtil;
using SHSchool.Data;
using Campus.DocumentValidator;
using FISCA.Presentation.Controls;
using FISCA.LogAgent;

namespace SHSchool.Behavior
{
    class ImportMoralScore : ImportWizard
    {
        //設定檔
        private ImportOption mOption;
        //Log內容
        private StringBuilder mstrLog = new StringBuilder();

        //學生已存在之日常生活表現記錄
        //private List<SHMoralScoreRecord> MoralScoreList;

        //學生Record,與學號對應
        private Dictionary<string, SHStudentRecord> StudentRecordByStudentNumber = new Dictionary<string, SHStudentRecord>();
        //文字描述代碼表
        private Dictionary<string, List<string>> faceDic;

        /// <summary>
        /// 準備動作
        /// </summary>
        public override void Prepare(ImportOption Option)
        {
            mOption = Option;
        }

        /// <summary>
        /// 依學生學號資料
        /// 取得學生的日常生活表現資料
        /// </summary>
        private List<SHMoralScoreRecord> GetMoralScoreList(List<Campus.DocumentValidator.IRowStream> Rows)
        {
            //學號清單
            List<string> StudentNumberList = new List<string>();
            foreach (IRowStream Row in Rows)
            {
                string StudentNumber = Row.GetValue("學號");
                if (!StudentNumberList.Contains(StudentNumber))
                {
                    StudentNumberList.Add(StudentNumber);
                }
            }
            //比對學號&取得學生Record
            List<SHStudentRecord> importStudRecList = new List<SHStudentRecord>();

            foreach (SHStudentRecord each in SHStudent.SelectAll())
            {
                if (each.Status == StudentRecord.StudentStatus.刪除) //狀態為刪除者排除
                    continue;

                if (StudentNumberList.Contains(each.StudentNumber)) //包含於匯入資料"學號清單之學生
                {
                    importStudRecList.Add(each);

                    //建立學號比對學生Record字典
                    if (!StudentRecordByStudentNumber.ContainsKey(each.StudentNumber))
                    {
                        StudentRecordByStudentNumber.Add(each.StudentNumber, each);
                    }
                }

            }

            //取得學生ID
            List<string> StudentIDList = importStudRecList.Select(x => x.ID).ToList();
            //一次取得學生日常生活表現清單
            return SHMoralScore.SelectByStudentIDs(StudentIDList);
        }

        /// <summary>
        /// 由學號與學年度學期,取得學生日常生活表現Record
        /// 如果沒有已存在Record,則是Null
        /// </summary>
        /// <param name="StudentNumber">學號</param>
        /// <param name="SchoolYear">學年度</param>
        /// <param name="Semester">學期</param>
        /// <returns></returns>
        public SHMoralScoreRecord GetMoralScoreRecord(List<SHMoralScoreRecord> MoralScoreList, string StudentNumber, int SchoolYear, int Semester)
        {
            foreach (SHMoralScoreRecord each in MoralScoreList)
            {
                if (each.Student.StudentNumber == StudentNumber && each.SchoolYear == SchoolYear && each.Semester == Semester)
                {
                    return each;
                }
            }
            return null;
        }

        /// <summary>
        /// 每1000筆資料,分批執行匯入
        /// Return是Log資訊
        /// </summary>
        public override string Import(List<Campus.DocumentValidator.IRowStream> Rows)
        {
            List<MoralScoreLogRobot> LogList = new List<MoralScoreLogRobot>();

            //取得匯入資料中,所有學生編號下的的日常生活表現資料
            List<SHMoralScoreRecord> MoralScoreList = GetMoralScoreList(Rows);

            List<SHMoralScoreRecord> InsertList = new List<SHMoralScoreRecord>();
            List<SHMoralScoreRecord> UpdateList = new List<SHMoralScoreRecord>();
            int NochangeCount = 0; //未處理資料記數

            if (mOption.Action == ImportAction.InsertOrUpdate)
            {
                foreach (IRowStream Row in Rows)
                {
                    int SchoolYear = int.Parse(Row.GetValue("學年度"));
                    int Semester = int.Parse(Row.GetValue("學期"));
                    string StudentNumber = Row.GetValue("學號");
                    SHStudentRecord student = StudentRecordByStudentNumber[StudentNumber];

                    SHMoralScoreRecord SSR = GetMoralScoreRecord(MoralScoreList, StudentNumber, SchoolYear, Semester);

                    //Log機器人
                    MoralScoreLogRobot log = new MoralScoreLogRobot(student, faceDic.Keys.ToList());

                    if (SSR == null) //新增SHMoralScoreRecord
                    {
                        SHMoralScoreRecord newRecord = new SHMoralScoreRecord();
                        newRecord.RefStudentID = student.ID;
                        newRecord.SchoolYear = SchoolYear;
                        newRecord.Semester = Semester;
                        newRecord.TextScore = GetInsertTextScore(Row);
                        //Log
                        log.SetOldLost();
                        log.SetNew(newRecord);

                        InsertList.Add(newRecord);
                    }
                    else //已存在資料需要修改 or 覆蓋
                    {
                        //Log(old)
                        log.SetOld(SSR);

                        //取得新的Xml結構
                        SSR.TextScore = GetUpdateTextScore(SSR, Row);

                        //Log(New)
                        log.SetNew(SSR);
                        if (!string.IsNullOrEmpty(log.LogToString()))
                        {
                            UpdateList.Add(SSR);
                        }
                    }
                    LogList.Add(log);
                }
            }
            if (InsertList.Count > 0)
            {
                try
                {
                    SHMoralScore.Insert(InsertList);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("於新增資料時發生錯誤!!\n" + ex.Message);
                }
            }
            if (UpdateList.Count > 0)
            {
                try
                {
                    SHMoralScore.Update(UpdateList);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("於更新資料時發生錯誤!!\n" + ex.Message);
                }
            }
            //批次記錄Log
            LogSaver Batch = FISCA.LogAgent.ApplicationLog.CreateLogSaverInstance();
            foreach (MoralScoreLogRobot each in LogList)
            {
                if (!string.IsNullOrEmpty(each.LogToString()))
                {
                    Batch.AddBatch("學務系統.匯入文字評量", "student", each._student.ID, each.LogToString());
                }
                else
                {
                    NochangeCount++;
                }
            }
            Batch.LogBatch(true);
            //Summary Log
            StringBuilder sbSummary = new StringBuilder();
            sbSummary.AppendLine("匯入文字評量操作：");
            sbSummary.AppendLine("新增：" + InsertList.Count + "筆資料");
            sbSummary.AppendLine("更新：" + UpdateList.Count + "筆資料");
            sbSummary.AppendLine("未更動：" + NochangeCount + "筆資料");
            FISCA.LogAgent.ApplicationLog.Log("學務系統.匯入文字評量", "匯入", sbSummary.ToString());

            //StringBuilder sbSummary2 = new StringBuilder();
            //sbSummary2.AppendLine("匯入文字評量操作：");
            //sbSummary2.AppendLine("新增：" + InsertList.Count + "筆資料");
            //sbSummary2.AppendLine("更新：" + UpdateList.Count + "筆資料");
            //sbSummary2.AppendLine("未更動：" + NochangeCount + "筆資料");

            return "";
        }

        #region 驗證規則
        /// <summary>
        /// 取得驗證規則(動態建置XML內容)
        /// </summary>
        public override string GetValidateRule()
        {
            //取得使用者的設定檔
            faceDic = GetConfig.GetWordCommentDic();

            //動態建立XmlRule
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(Properties.Resources.ImportMoralScoreValRule);
            makeXmlContent(xmlDoc);
            return xmlDoc.InnerXml;
        }

        /// <summary>
        /// 依使用者設定檔"文字評量代碼表"
        /// 建置動態的XML規則
        /// </summary>
        private void makeXmlContent(XmlDocument xmlDoc)
        {
            //1. get the FiledList Element
            XmlElement elmFieldList = (XmlElement)xmlDoc.DocumentElement.SelectSingleNode("FieldList");

            //2. Append each field element to FieldList Element.
            foreach (string fieldName in faceDic.Keys)
            {
                XmlElement elmField = xmlDoc.CreateElement("Field");
                elmField.SetAttribute("Required", "False");
                elmField.SetAttribute("Name", fieldName);
                elmField.SetAttribute("EmptyAlsoValidate", "False");
                elmField.SetAttribute("Description", "文字評量代碼表");

                //XmlElement elmValidate = xmlDoc.CreateElement("Validate");
                //elmValidate.SetAttribute("AutoCorrect", "False");
                //elmValidate.SetAttribute("Description", "「" + fieldName + "」與代碼表內容不一致");
                //elmValidate.SetAttribute("ErrorType", "Warning");
                //elmValidate.SetAttribute("Validator", "文字評量代碼表列舉(" + fieldName + ")");
                //elmValidate.SetAttribute("When", "");
                //elmField.AppendChild(elmValidate);

                elmFieldList.AppendChild(elmField);
            }

            //3. Append each Validator element to ValidatorList Element.
            //elmFieldList = (XmlElement)xmlDoc.DocumentElement.SelectSingleNode("ValidatorList");
            //foreach (string fieldName in faceDic.Keys)
            //{
            //    XmlElement elmField = xmlDoc.CreateElement("FieldValidator");
            //    elmField.SetAttribute("Name", "文字評量代碼表列舉(" + fieldName + ")");
            //    elmField.SetAttribute("Type", "Enumeration");
            //    foreach (string item in faceDic[fieldName])
            //    {
            //        XmlElement elmValidate = xmlDoc.CreateElement("Item");
            //        elmValidate.SetAttribute("Value", item);
            //        elmField.AppendChild(elmValidate);
            //    }
            //    elmFieldList.AppendChild(elmField);
            //}

        }

        #endregion

        /// <summary>
        /// 傳入Row,以建立XML資料內容
        /// </summary>
        private XmlElement GetInsertTextScore(IRowStream Row)
        {
            XmlDocument XmlDoc = new XmlDocument();
            XmlElement XmlElementB1 = XmlDoc.CreateElement("Content");
            foreach (string each in faceDic.Keys)
            {
                if (mOption.SelectedFields.Contains(each)) //如果使用者已選欄位
                {
                    XmlElement elmMorality = XmlDoc.CreateElement("Morality");
                    elmMorality.SetAttribute("Face", each);
                    elmMorality.InnerText = Row.GetValue(each);
                    XmlElementB1.AppendChild(elmMorality);
                }
                else //如果使用者未選該欄位
                {
                    XmlElement elmMorality = XmlDoc.CreateElement("Morality");
                    elmMorality.SetAttribute("Face", each);
                    elmMorality.InnerText = ""; //未選擇欄位,將會填入空值
                    XmlElementB1.AppendChild(elmMorality);
                }
            }
            XmlDoc.AppendChild(XmlElementB1);
            return XmlDoc.DocumentElement;
        }

        /// <summary>
        /// 傳入MoralScore與Row,以建立資料狀態
        /// </summary>
        private XmlElement GetUpdateTextScore(SHMoralScoreRecord SSR, IRowStream Row)
        {
            XmlDocument XmlDoc = new XmlDocument();

            #region TextScores內容為空,則建立一個預設根節點

            if (SSR.TextScore != null)
            {
                if (!string.IsNullOrEmpty(SSR.TextScore.OuterXml))
                {
                    XmlDoc.LoadXml(SSR.TextScore.OuterXml);
                }
                else
                {
                    XmlDoc.LoadXml("<Content/>");
                }
            }
            else
            {
                XmlDoc.LoadXml("<Content/>");
            }

            #endregion

            XmlElement XmlElementB1;

            #region 把Content置換成TextScore

            if (XmlDoc.DocumentElement.Name == "Content")
            {
                XmlElementB1 = (XmlElement)XmlDoc.SelectSingleNode("Content");
            }
            else if (XmlDoc.DocumentElement.Name == "TextScore")
            {
                XmlDocument copy = new XmlDocument();
                copy.LoadXml("<Content/>");

                XmlElement Xmlaa = (XmlElement)XmlDoc.SelectSingleNode("TextScore");
                foreach (XmlElement Xmlbb in Xmlaa.SelectNodes("Morality"))
                {
                    XmlElement cc = copy.CreateElement("Morality");
                    cc.SetAttribute("Face", Xmlbb.GetAttribute("Face"));
                    cc.InnerText = Xmlbb.InnerText;
                    copy.DocumentElement.AppendChild(cc);
                }

                XmlDoc = copy;
                XmlElementB1 = (XmlElement)XmlDoc.DocumentElement;//.SelectSingleNode("Content");
            }
            else
            {
                XmlElementB1 = XmlDoc.DocumentElement;
            } 

            #endregion

            //以設定檔為主,掃瞄Xml內容
            foreach (string each in faceDic.Keys)
            {
                if (mOption.SelectedFields.Contains(each)) //如果使用者選擇該欄位
                {
                    XmlElement elmMorality = (XmlElement)XmlElementB1.SelectSingleNode("Morality[@Face='" + each + "']");
                    if (elmMorality != null) //包含此內容
                    {
                        if (elmMorality.GetAttribute("Face") == each)
                        {
                            elmMorality.InnerText = Row.GetValue(each);
                        }
                    }
                    else //如果Xml內不包含此資料欄位,則新增於Xml內
                    {
                        elmMorality = (XmlElement)XmlDoc.CreateElement("Morality");
                        elmMorality.SetAttribute("Face", each);
                        elmMorality.InnerText = Row.GetValue(each);
                        XmlElementB1.AppendChild(elmMorality);
                    }
                }
                else //使用者未選擇該欄位,則不進行更動
                {
                    //....
                }
            }
            return XmlDoc.DocumentElement;
        }

        /// <summary>
        /// 設定匯入功能,所提供的匯入動作
        /// </summary>
        public override ImportAction GetSupportActions()
        {
            //新增或更新
            return ImportAction.InsertOrUpdate;
        }
    }
}
