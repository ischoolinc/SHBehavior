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
    class ImportComment : ImportWizard
    {
        //設定檔
        private ImportOption mOption;
        //Log內容
        private StringBuilder mstrLog = new StringBuilder();

        //學生已存在之日常生活表現記錄
        //private List<SHMoralScoreRecord> MoralScoreList;

        //學生Record,與學號對應
        private Dictionary<string, SHStudentRecord> StudentRecordByStudentNumber = new Dictionary<string, SHStudentRecord>();

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
            List<CommentLogRobot> LogList = new List<CommentLogRobot>();

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
                    string Comment = Row.GetValue("導師評語");
                    SHStudentRecord student = StudentRecordByStudentNumber[StudentNumber];
                    SHMoralScoreRecord SSR = GetMoralScoreRecord(MoralScoreList, StudentNumber, SchoolYear, Semester);

                    //Log機器人
                    CommentLogRobot log = new CommentLogRobot(student);

                    if (SSR == null) //新增SHMoralScoreRecord
                    {
                        SHMoralScoreRecord newRecord = new SHMoralScoreRecord();
                        newRecord.RefStudentID = student.ID;
                        newRecord.SchoolYear = SchoolYear;
                        newRecord.Semester = Semester;
                        newRecord.Comment = Comment;
                        //Log
                        log.SetOldLost();
                        log.SetNew(newRecord);
                        InsertList.Add(newRecord);
                    }
                    else //已存在資料需要修改 or 覆蓋
                    {
                        //Log(old)
                        log.SetOld(SSR);
                        SSR.Comment = Comment;

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
                catch(Exception ex)
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
            foreach (CommentLogRobot each in LogList)
            {
                if (!string.IsNullOrEmpty(each.LogToString()))
                {
                    Batch.AddBatch("學務系統.匯入導師評語", "student", each._student.ID, each.LogToString());
                }
                else
                {
                    NochangeCount++;
                }
            }
            Batch.LogBatch(true);
            //Summary Log
            StringBuilder sbSummary = new StringBuilder();
            sbSummary.AppendLine("匯入導師評語操作：");
            sbSummary.AppendLine("新增：" + InsertList.Count + "筆資料");
            sbSummary.AppendLine("更新：" + UpdateList.Count + "筆資料");
            sbSummary.AppendLine("未更動：" + NochangeCount + "筆資料");
            FISCA.LogAgent.ApplicationLog.Log("學務系統.匯入導師評語", "匯入", sbSummary.ToString());

            //StringBuilder sbSummary2 = new StringBuilder();
            //sbSummary2.AppendLine("匯入導師評語：");
            //sbSummary2.AppendLine("新增：" + InsertList.Count + "筆資料");
            //sbSummary2.AppendLine("更新：" + UpdateList.Count + "筆資料");
            //sbSummary2.AppendLine("未更動：" + NochangeCount + "筆資料");

            return "";
        }

        /// <summary>
        /// 取得驗證規則(動態建置XML內容)
        /// </summary>
        public override string GetValidateRule()
        {
            //動態建立XmlRule
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(Properties.Resources.ImportCommentValRule);
            return xmlDoc.InnerXml;
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
