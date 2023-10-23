using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using FISCA.DSAUtil;
using K12.Data;
using K12.BusinessLogic;
using SHSchool.Data;
using System.Data;

namespace SHSchool.DailyManifestation
{
    class StudentDataObj
    {
        private List<string> _BehaviorList;

        public StudentDataObj(List<string> BehaviorList)
        {
            _BehaviorList = BehaviorList;
        }
        #region 物件資料欄位

        /// <summary>
        /// 學生基本資料物件
        /// </summary>
        public StudentRecord StudentRecord { get; set; }

        /// <summary>
        /// 戶籍電話
        /// </summary>
        public string PhonePermanent { get; set; }

        /// <summary>
        /// 聯絡電話
        /// </summary>
        public string PhoneContact { get; set; }

        /// <summary>
        /// 監護人姓名
        /// </summary>
        public string CustodianName { get; set; }

        /// <summary>
        /// 監護人稱謂
        /// </summary>
        public string CustodianTitle { get; set; }

        /// <summary>
        /// 戶籍地址
        /// </summary>
        public string AddressPermanent { get; set; }

        /// <summary>
        /// 聯絡地址
        /// </summary>
        public string AddressMailing { get; set; }

        /// <summary>
        /// 學期歷程
        /// </summary>
        public SemesterHistoryRecord SemesterHistory { get; set; }

        /// <summary>
        /// 獎勵資料清單
        /// </summary>
        public List<MeritRecord> ListMerit = new List<MeritRecord>();

        /// <summary>
        /// 懲戒資料清單
        /// </summary>
        public List<DemeritRecord> ListDeMerit = new List<DemeritRecord>();

        /// <summary>
        /// 日常生活表現
        /// </summary>
        public List<SHMoralScoreRecord> ListMoralScore = new List<SHMoralScoreRecord>();

        /// <summary>
        /// 異動記錄
        /// </summary>
        public List<UpdateRecordRecord> ListUpdateRecord = new List<UpdateRecordRecord>();

        /// <summary>
        /// 自動統計缺曠獎懲
        /// </summary>
        public List<AutoSummaryRecord> ListAutoSummary = new List<AutoSummaryRecord>();

        /// <summary>
        /// 日常生活表現,學年度學期,標題:內容
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> TextScoreDic = new Dictionary<string, Dictionary<string, string>>();

        /// <summary>
        /// 統計資料,學年度學期,假別:內容
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> DicAsbs = new Dictionary<string, Dictionary<string, string>>();

        /// <summary>
        /// 獎懲統計資料
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> DicMeritDemerit = new Dictionary<string, Dictionary<string, string>>();

        /// <summary>
        /// 社團活動
        /// </summary>
        public Dictionary<string, string> AssnDic = new Dictionary<string, string>();

        /// <summary>
        /// 努力程度物件
        /// </summary>
        //private EffortMapper Effor = new EffortMapper();

        //班級座號
        public string GradeYear11 { get; set; }
        public string GradeYear12 { get; set; }
        public string GradeYear21 { get; set; }
        public string GradeYear22 { get; set; }
        public string GradeYear31 { get; set; }
        public string GradeYear32 { get; set; }

        //應到日數
        public Dictionary<string, string> SchoolDay = new Dictionary<string, string>();

        #endregion

        Dictionary<string, List<string>> _merit_flagDic;

        /// <summary>
        /// 初始化資料內容
        /// </summary>
        public void SetupData(Dictionary<string, List<string>> merit_flagDic)
        {
            _merit_flagDic = merit_flagDic;
            //取得新生資料
            NewStudent();

            //資料排序
            DataSort();

            //處理學期歷程
            ExecuteSemesterSet();

            //處理社團活動
            //ExecuteAssnCode();

            //處理日常生活表現
            ExecuteMoralScore();

        }

        #region 新生異動

        /// <summary>
        /// 畢業國小
        /// </summary>
        public string UpdataGraduateSchool { get; set; }

        /// <summary>
        /// 入學核准日期
        /// </summary>
        public string UpdataADDate { get; set; }

        /// <summary>
        /// 入學核准文號
        /// </summary>
        public string UpdataADNumber { get; set; }

        /// <summary>
        /// 取得新生入學資訊
        /// </summary>
        private void NewStudent()
        {
            foreach (UpdateRecordRecord each in ListUpdateRecord)
            {
                if (each.UpdateCode == "1")
                {
                    //UpdataGraduateSchool = each.GraduateSchool;
                    UpdataADDate = each.ADDate;
                    UpdataADNumber = each.ADNumber;
                }
            }
        }

        #endregion

        #region 排序

        private void DataSort()
        {
            ListMerit.Sort(new Comparison<MeritRecord>(MeritSort));
            ListDeMerit.Sort(new Comparison<DemeritRecord>(DeMeritSort));
            ListUpdateRecord.Sort(new Comparison<UpdateRecordRecord>(UpdateRecordSort));
            ListMoralScore.Sort(new Comparison<MoralScoreRecord>(MoralScoreSort));

        }

        private int MeritSort(MeritRecord x, MeritRecord y)
        {
            return x.OccurDate.CompareTo(y.OccurDate);
        }

        private int DeMeritSort(DemeritRecord x, DemeritRecord y)
        {
            return x.OccurDate.CompareTo(y.OccurDate);
        }

        private int UpdateRecordSort(UpdateRecordRecord x, UpdateRecordRecord y)
        {
            DateTime dt1;
            DateTime.TryParse(x.UpdateDate, out dt1);
            DateTime dt2;
            DateTime.TryParse(y.UpdateDate, out dt2);
            return dt1.CompareTo(dt2);
        }

        private int MoralScoreSort(MoralScoreRecord x, MoralScoreRecord y)
        {
            string xx = x.SchoolYear.ToString() + x.Semester.ToString();
            string yy = y.SchoolYear.ToString() + y.Semester.ToString();
            return xx.CompareTo(yy);
        }

        #endregion

        #region 學期歷程

        private int SortMoralScore(MoralScoreRecord x1, MoralScoreRecord x2)
        {
            string schoolyear1 = x1.SchoolYear.ToString().PadLeft(5, '0');
            schoolyear1 += x1.Semester.ToString().PadLeft(5, '0');

            string schoolyear2 = x2.SchoolYear.ToString().PadLeft(5, '0');
            schoolyear2 += x2.Semester.ToString().PadLeft(5, '0');

            return schoolyear1.CompareTo(schoolyear2);

        }

        private int SortAutoSummary(AutoSummaryRecord x1, AutoSummaryRecord x2)
        {
            string schoolyear1 = x1.SchoolYear.ToString().PadLeft(5, '0');
            schoolyear1 += x1.Semester.ToString().PadLeft(5, '0');

            string schoolyear2 = x2.SchoolYear.ToString().PadLeft(5, '0');
            schoolyear2 += x2.Semester.ToString().PadLeft(5, '0');

            return schoolyear1.CompareTo(schoolyear2);

        }
        //處理學期歷程資料
        private void ExecuteSemesterSet()
        {
            List<K12.Data.SemesterHistoryItem> list = SemesterHistory.SemesterHistoryItems;

            //學期歷程重讀判斷
            list = ReSetHistory(list);

            foreach (K12.Data.SemesterHistoryItem each in list)
            {
                string ClassNameAndSeatNo = each.SchoolYear.ToString() + "/" + each.Semester;

                if (each.GradeYear == 1 || each.GradeYear == 7)
                {
                    if (each.Semester == 1)
                    {
                        GradeYear11 = ClassNameAndSeatNo;
                    }
                    else
                    {
                        GradeYear12 = ClassNameAndSeatNo;
                    }
                }
                else if (each.GradeYear == 2 || each.GradeYear == 8)
                {
                    if (each.Semester == 1)
                    {
                        GradeYear21 = ClassNameAndSeatNo;
                    }
                    else
                    {
                        GradeYear22 = ClassNameAndSeatNo;
                    }
                }
                else if (each.GradeYear == 3 || each.GradeYear == 9)
                {
                    if (each.Semester == 1)
                    {
                        GradeYear31 = ClassNameAndSeatNo;
                    }
                    else
                    {
                        GradeYear32 = ClassNameAndSeatNo;
                    }
                }
            }
        }

        private int SortSchoolYear(SemesterHistoryItem a, SemesterHistoryItem b)
        {
            string stringA = a.SchoolYear.ToString().PadLeft(5, '0') + a.Semester.ToString().PadLeft(5, '0');
            string stringB = b.SchoolYear.ToString().PadLeft(5, '0') + b.Semester.ToString().PadLeft(5, '0');

            return stringA.CompareTo(stringB);
        }

        /// <summary>
        /// 判斷並移除重讀之學期歷程
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private List<SemesterHistoryItem> ReSetHistory(List<SemesterHistoryItem> list)
        {
            Dictionary<string, SemesterHistoryItem> _dic = new Dictionary<string, SemesterHistoryItem>();

            //進行資料排續
            list.Sort(SortSchoolYear);
            // 0009900001
            // 0009900002
            // 0010000001
            // 0010000002

            //98 1 
            //99 1
            //99 2
            //100 1
            //100 2


            //透過迴圈進行資料填入 / 利用資料覆蓋方式來取最新值
            //1 1 -> 98 1 -> 99 1
            //1 2
            //2 1
            //2 2
            //3 1
            //3 2

            //99學年度1學期
            //100學年度1學期 <--取大的學年度
            foreach (SemesterHistoryItem each in list)
            {
                string GradeYearSemester = each.GradeYear.ToString().PadLeft(5, '0') + each.Semester.ToString().PadLeft(5, '0');
                if (!_dic.ContainsKey(GradeYearSemester))
                {
                    _dic.Add(GradeYearSemester, each);
                }
                else //重覆就覆蓋
                {
                    _dic[GradeYearSemester] = each;
                }
            }

            List<SemesterHistoryItem> _list = new List<SemesterHistoryItem>();
            foreach (string each in _dic.Keys)
            {
                _list.Add(_dic[each]);
            }

            return _list;
        }

        #endregion

        /// <summary>
        /// 日常生活表現資料整理
        /// </summary>
        private void ExecuteMoralScore()
        {
            ListAutoSummary.Sort(SortAutoSummary);

            foreach (AutoSummaryRecord auto in ListAutoSummary)
            {
                #region 處理缺曠
                //學年度學期
                string SchoolYearSemester = auto.SchoolYear.ToString() + "/" + auto.Semester.ToString();

                DSXmlHelper helper2 = new DSXmlHelper(auto.AutoSummary);

                //如果有缺曠
                if (helper2.GetElement("AttendanceStatistics") != null)
                {
                    if (helper2.GetElements("AttendanceStatistics/Absence").Count() == 0)
                        continue;

                    if (!DicAsbs.ContainsKey(SchoolYearSemester))
                    {
                        DicAsbs.Add(SchoolYearSemester, new Dictionary<string, string>());

                        foreach (XmlElement item in helper2.GetElements("AttendanceStatistics/Absence"))
                        {
                            DicAsbs[SchoolYearSemester].Add(item.GetAttribute("Name") + item.GetAttribute("PeriodType"), item.GetAttribute("Count"));
                        }
                    }
                }
                #endregion
            }

            foreach (AutoSummaryRecord auto in ListAutoSummary)
            {
                #region 處理獎懲
                //學年度學期
                string SchoolYearSemester = auto.SchoolYear.ToString() + "/" + auto.Semester.ToString();

                DSXmlHelper helper2 = new DSXmlHelper(auto.AutoSummary);

                int Count = CountMeritDemerit(helper2);



                //如果有獎懲
                if (Count != 0)
                {
                    if (!DicMeritDemerit.ContainsKey(SchoolYearSemester))
                    {
                        DicMeritDemerit.Add(SchoolYearSemester, new Dictionary<string, string>());

                        if (helper2.GetElement("DisciplineStatistics/Merit") != null)
                        {
                            DicMeritDemerit[SchoolYearSemester].Add("大功", helper2.GetAttribute("DisciplineStatistics/Merit/@A"));
                            DicMeritDemerit[SchoolYearSemester].Add("小功", helper2.GetAttribute("DisciplineStatistics/Merit/@B"));
                            DicMeritDemerit[SchoolYearSemester].Add("嘉獎", helper2.GetAttribute("DisciplineStatistics/Merit/@C"));
                        }

                        if (helper2.GetElement("DisciplineStatistics/Demerit") != null)
                        {
                            DicMeritDemerit[SchoolYearSemester].Add("大過", helper2.GetAttribute("DisciplineStatistics/Demerit/@A"));
                            DicMeritDemerit[SchoolYearSemester].Add("小過", helper2.GetAttribute("DisciplineStatistics/Demerit/@B"));
                            DicMeritDemerit[SchoolYearSemester].Add("警告", helper2.GetAttribute("DisciplineStatistics/Demerit/@C"));
                        }
                    }
                }

                if (_merit_flagDic.ContainsKey(auto.RefStudentID))
                {
                    if (_merit_flagDic[auto.RefStudentID].Contains(SchoolYearSemester))
                    {
                        if (!DicMeritDemerit.ContainsKey(SchoolYearSemester))
                        {
                            DicMeritDemerit.Add(SchoolYearSemester, new Dictionary<string, string>());
                        }

                        DicMeritDemerit[SchoolYearSemester].Add("留察", "是");

                    }
                }



                #endregion
            }

            ListMoralScore.Sort(SortMoralScore);

            foreach (SHMoralScoreRecord each in ListMoralScore)
            {
                //學年度學期
                string SchoolYearSemester = each.SchoolYear.ToString() + "/" + each.Semester.ToString();

                #region 日常生活表現
                if (!TextScoreDic.ContainsKey(SchoolYearSemester))
                {
                    TextScoreDic.Add(SchoolYearSemester, new Dictionary<string, string>());
                }

                //增加一格空白,用來把資料跟日常生活表現上的導師評語區分
                if (!string.IsNullOrEmpty(each.Comment))
                    TextScoreDic[SchoolYearSemester].Add("導師評語 ", each.Comment);

                DSXmlHelper helper1 = new DSXmlHelper(each.TextScore);

                foreach (XmlElement each1 in helper1.BaseElement.SelectNodes("Morality"))
                {
                    string strFace = each1.GetAttribute("Face");
                    if (_BehaviorList.Contains(strFace))
                    {
                        if (!TextScoreDic[SchoolYearSemester].ContainsKey(strFace))
                        {
                            TextScoreDic[SchoolYearSemester].Add(strFace, each1.InnerText);
                        }
                    }
                }
                #endregion

            }
        }

        private int CountMeritDemerit(DSXmlHelper xmlElement)
        {
            int MA = IntCheckPas(xmlElement.GetAttribute("DisciplineStatistics/Merit/@A"));
            int MB = IntCheckPas(xmlElement.GetAttribute("DisciplineStatistics/Merit/@B"));
            int MC = IntCheckPas(xmlElement.GetAttribute("DisciplineStatistics/Merit/@C"));
            int DemA = IntCheckPas(xmlElement.GetAttribute("DisciplineStatistics/Demerit/@A"));
            int DemB = IntCheckPas(xmlElement.GetAttribute("DisciplineStatistics/Demerit/@B"));
            int DemC = IntCheckPas(xmlElement.GetAttribute("DisciplineStatistics/Demerit/@C"));
            return MA + MB + MC + DemA + DemB + DemC;
        }

        private int IntCheckPas(string a)
        {
            int x = 0;
            int.TryParse(a, out x);
            return x;
        }
    }
}
