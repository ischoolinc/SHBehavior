using K12.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace 日常生活表現預警表
{
    /// <summary>
    /// 一個學生
    /// </summary>
    class StudRecord
    {
        int _school_year;
        int _semester;

        int sumSection = 0;

        //功過換算
        MeritDemeritReduceRecord _reduceRecord;
        int MeritAtoB = 0;
        int MeritBtoC = 0;
        int DemeritAtoB = 0;
        int DemeritBtoC = 0;

        public StudRecord(DataRow row, string school_year, string semester, MeritDemeritReduceRecord reduceRecord)
        {
            _reduceRecord = reduceRecord;

            if (_reduceRecord.MeritAToMeritB.HasValue)
                MeritAtoB = _reduceRecord.MeritAToMeritB.Value;

            if (_reduceRecord.MeritBToMeritC.HasValue)
                MeritBtoC = _reduceRecord.MeritBToMeritC.Value;

            if (_reduceRecord.DemeritAToDemeritB.HasValue)
                DemeritAtoB = _reduceRecord.DemeritAToDemeritB.Value;

            if (_reduceRecord.DemeritBToDemeritC.HasValue)
                DemeritBtoC = _reduceRecord.DemeritBToDemeritC.Value;

            int.TryParse(school_year, out _school_year);
            int.TryParse(semester, out _semester);

            StudentID = "" + row["student_id"];
            StudentName = "" + row["student_name"];
            SeatNo = "" + row["seat_no"];
            StudentNumber = "" + row["student_number"];

            ClassID = "" + row["class_id"];
            ClassName = "" + row["class_name"];
            ClassNumber = "" + row["display_order"];
            GradeYear = "" + row["grade_year"];

            TeacherID = "" + row["teacher_id"];


            AttendanceList = new List<AttendanceRecord>();
            MeritList = new List<MeritRecord>();
            DemeritList = new List<DemeritRecord>();
        }

        //基本資料

        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 學生姓名
        /// </summary>
        public string StudentName { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        public string SeatNo { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        public string StudentNumber { get; set; }

        /// <summary>
        /// 班級系統編號
        /// </summary>
        public string ClassID { get; set; }

        /// <summary>
        /// 班級名稱
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 班級序號
        /// </summary>
        public string ClassNumber { get; set; }

        public string TeacherID { get; set; }

        /// <summary>
        /// 年級
        /// </summary>
        public string GradeYear { get; set; }


        public int 本期大功 { get; set; }
        public int 本期小功 { get; set; }
        public int 本期嘉獎 { get; set; }
        public int 本期大過 { get; set; }
        public int 本期小過 { get; set; }
        public int 本期警告 { get; set; }

        //2020/6/8 - 歷年資料不予換算
        public int 歷年大功 { get; set; }
        public int 歷年小功 { get; set; }
        public int 歷年嘉獎 { get; set; }
        public int 歷年大過 { get; set; }
        public int 歷年小過 { get; set; }
        public int 歷年警告 { get; set; }

        public int 歷年功過換算大功 { get; set; }
        public int 歷年功過換算小功 { get; set; }
        public int 歷年功過換算嘉獎 { get; set; }
        public int 歷年功過換算大過 { get; set; }
        public int 歷年功過換算小過 { get; set; }
        public int 歷年功過換算警告 { get; set; }

        public int 歷年銷過大過 { get; set; }
        public int 歷年銷過小過 { get; set; }
        public int 歷年銷過警告 { get; set; }

        //換算概念
        //1.全部換算為最小單位
        //2.進行相減
        //3.換算回相對單位
        public int 歷年功過相抵大功 { get; set; }
        public int 歷年功過相抵小功 { get; set; }
        public int 歷年功過相抵嘉獎 { get; set; }
        public int 歷年功過相抵大過 { get; set; }
        public int 歷年功過相抵小過 { get; set; }
        public int 歷年功過相抵警告 { get; set; }

        public bool 是否留查 = false;
        public bool 是否全勤 = true;

        public List<string> RemarkList = new List<string>();

        //功過相抵
        int 獎勵 = 0;
        int 懲戒 = 0;

        /// <summary>
        /// 缺曠資料
        /// </summary>
        public List<AttendanceRecord> AttendanceList { get; set; }

        /// <summary>
        /// 缺曠
        /// </summary>
        public Dictionary<string, int> AttendanceDic { get; set; }

        /// <summary>
        /// 獎勵資料
        /// </summary>
        public List<MeritRecord> MeritList { get; set; }

        /// <summary>
        /// 懲戒資料
        /// </summary>
        public List<DemeritRecord> DemeritList { get; set; }

        /// <summary>
        /// 分類獎懲
        /// </summary>
        public void RunDeMerit(int disciplinaryMeasures)
        {
            //獎勵
            foreach (MeritRecord meric in MeritList)
            {
                if (_school_year == meric.SchoolYear && _semester == meric.Semester)
                {
                    //本期資料
                    本期大功 += meric.MeritA.HasValue ? meric.MeritA.Value : 0;
                    本期小功 += meric.MeritB.HasValue ? meric.MeritB.Value : 0;
                    本期嘉獎 += meric.MeritC.HasValue ? meric.MeritC.Value : 0;
                }

                //非本期資料
                歷年大功 += meric.MeritA.HasValue ? meric.MeritA.Value : 0;
                歷年小功 += meric.MeritB.HasValue ? meric.MeritB.Value : 0;
                歷年嘉獎 += meric.MeritC.HasValue ? meric.MeritC.Value : 0;
            }

            //懲戒
            foreach (DemeritRecord demeric in DemeritList)
            {
                //判斷本期是否有留察
                if (_school_year == demeric.SchoolYear && _semester == demeric.Semester)
                {
                    if (demeric.MeritFlag == "2")
                    {
                        是否留查 = true;
                    }
                }

                if (demeric.MeritFlag == "0")
                {
                    if (demeric.Cleared == "")
                    {
                        if (_school_year == demeric.SchoolYear && _semester == demeric.Semester)
                        {
                            //本期資料
                            本期大過 += demeric.DemeritA.HasValue ? demeric.DemeritA.Value : 0;
                            本期小過 += demeric.DemeritB.HasValue ? demeric.DemeritB.Value : 0;
                            本期警告 += demeric.DemeritC.HasValue ? demeric.DemeritC.Value : 0;
                        }

                        歷年大過 += demeric.DemeritA.HasValue ? demeric.DemeritA.Value : 0;
                        歷年小過 += demeric.DemeritB.HasValue ? demeric.DemeritB.Value : 0;
                        歷年警告 += demeric.DemeritC.HasValue ? demeric.DemeritC.Value : 0;
                    }
                    else if (demeric.Cleared == "是")
                    {
                        歷年銷過大過 += demeric.DemeritA.HasValue ? demeric.DemeritA.Value : 0;
                        歷年銷過小過 += demeric.DemeritB.HasValue ? demeric.DemeritB.Value : 0;
                        歷年銷過警告 += demeric.DemeritC.HasValue ? demeric.DemeritC.Value : 0;
                    }
                }
            }
        }

        /// <summary>
        /// 分類缺曠資料
        /// </summary>
        public void RunAttendance(Dictionary<string, string> periodDic, Dictionary<string, List<string>> 顯示節次Dic, Dictionary<string, List<string>> 採記節次Dic)
        {
            AttendanceDic = new Dictionary<string, int>();
            foreach (AttendanceRecord attendace in AttendanceList)
            {
                if (_school_year == attendace.SchoolYear && _semester == attendace.Semester)
                {
                    foreach (AttendancePeriod period in attendace.PeriodDetail)
                    {
                        //查詢節次類別使用
                        if (periodDic.ContainsKey(period.Period))
                        {
                            string name = periodDic[period.Period] + period.AbsenceType;

                            //是否為選擇的節次類別
                            if (顯示節次Dic.ContainsKey(periodDic[period.Period]))
                            {
                                //是否為選擇的缺曠別
                                if (顯示節次Dic[periodDic[period.Period]].Contains(period.AbsenceType))
                                {
                                    //本缺曠別資料不存在
                                    if (!AttendanceDic.ContainsKey(name))
                                    {
                                        AttendanceDic.Add(name, 0);
                                    }
                                    AttendanceDic[name]++;
                                }
                            }
                        }

                        if (periodDic.ContainsKey(period.Period))
                        {
                            string name = periodDic[period.Period] + period.AbsenceType;

                            //是否為選擇的節次類別
                            if (採記節次Dic.ContainsKey(periodDic[period.Period]))
                            {
                                //是否為選擇的缺曠別
                                if (採記節次Dic[periodDic[period.Period]].Contains(period.AbsenceType))
                                {
                                    sumSection++;
                                    是否全勤 = false;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 進行功過換算作業
        /// </summary>
        public void RunOffset()
        {
            foreach (MeritRecord meric in MeritList)
            {
                int 大功換算 = 0;
                int 小功換算 = 0;
                int 嘉獎換算 = meric.MeritC.HasValue ? meric.MeritC.Value : 0;
                if (meric.MeritA.HasValue && meric.MeritA.Value != 0)
                {
                    大功換算 = meric.MeritA.Value * DemeritAtoB * DemeritBtoC;
                }

                if (meric.MeritB.HasValue && meric.MeritB.Value != 0)
                    小功換算 = meric.MeritB.Value * DemeritBtoC;

                獎勵 += 嘉獎換算 + 小功換算 + 大功換算;
            }

            foreach (DemeritRecord demeric in DemeritList)
            {
                //資料必須為未銷過
                if (demeric.Cleared == "")
                {
                    int 大過換算 = 0;
                    int 小過換算 = 0;
                    int 警告換算 = demeric.DemeritC.HasValue ? demeric.DemeritC.Value : 0;

                    if (demeric.DemeritA.HasValue && demeric.DemeritA.Value != 0)
                    {
                        大過換算 = demeric.DemeritA.Value * DemeritAtoB * DemeritBtoC;
                    }

                    if (demeric.DemeritB.HasValue && demeric.DemeritB.Value != 0)
                        小過換算 = demeric.DemeritB.Value * DemeritBtoC;

                    懲戒 += 警告換算 + 小過換算 + 大過換算;
                }
            }

            if (MeritList.Count > 0)
            {
                Crett ccc = cha(MeritAtoB, MeritBtoC, 獎勵);
                歷年功過換算大功 = ccc.歷年功過相抵A;
                歷年功過換算小功 = ccc.歷年功過相抵B;
                歷年功過換算嘉獎 = ccc.歷年功過相抵C;
            }

            if (DemeritList.Count > 0)
            {
                Crett ccc = cha(MeritAtoB, MeritBtoC, 懲戒);
                歷年功過換算大過 = ccc.歷年功過相抵A;
                歷年功過換算小過 = ccc.歷年功過相抵B;
                歷年功過換算警告 = ccc.歷年功過相抵C;
            }
        }

        /// <summary>
        /// 進行功過相抵作業
        /// </summary>
        public void RunOffsetAtoB()
        {
            if (獎勵 > 懲戒)
            {
                //餘獎勵
                int 獎懲功過相抵和 = 獎勵 - 懲戒;
                Crett ccc = cha(MeritAtoB, MeritBtoC, 獎懲功過相抵和);
                歷年功過相抵大功 = ccc.歷年功過相抵A;
                歷年功過相抵小功 = ccc.歷年功過相抵B;
                歷年功過相抵嘉獎 = ccc.歷年功過相抵C;
            }
            else if (獎勵 < 懲戒)
            {
                //餘懲戒
                int 獎懲功過相抵和 = 懲戒 - 獎勵;
                Crett ccc = cha(DemeritAtoB, DemeritBtoC, 獎懲功過相抵和);
                歷年功過相抵大過 = ccc.歷年功過相抵A;
                歷年功過相抵小過 = ccc.歷年功過相抵B;
                歷年功過相抵警告 = ccc.歷年功過相抵C;
            }
            else
            {
                //相抵後為0
            }


        }

        public Crett cha(int ab, int bc, int 總和)
        {
            Crett ccc = new Crett();
            int 比值A = ab * bc;
            int 比值B = ab;

            ccc.歷年功過相抵A = 總和 / 比值A;

            //餘數是否大於0
            //大於0則是嘉獎資料
            if (總和 % 比值A > 0)
            {
                //
                int d = 總和 % 比值A;
                if (d >= 比值B)
                {
                    ccc.歷年功過相抵B = d / 比值B;
                    ccc.歷年功過相抵C = d % 比值B;
                }
                else
                {
                    ccc.歷年功過相抵C = d;
                }
            }
            return ccc;
        }

        /// <summary>
        /// 是否列印
        /// </summary>
        public bool CheckSection(int _MiningSection, int _DisciplinaryMeasures, bool _IsAtoB相抵, bool _IsAtoB換算)
        {
            bool check = false;
            if (sumSection >= _MiningSection)
            {
                check = true;
                RemarkList.Add(string.Format("曠課滿{0}節", _MiningSection.ToString()));
            }

            if (_IsAtoB相抵)
            {
                if (歷年功過相抵大過 >= _DisciplinaryMeasures)
                {

                    check = true;
                    RemarkList.Add(string.Format("滿{0}大過", _DisciplinaryMeasures.ToString()));
                }
            }
            else
            {
                if (_IsAtoB換算)
                {
                    if (歷年功過換算大過 >= _DisciplinaryMeasures)
                    {

                        check = true;
                        RemarkList.Add(string.Format("滿{0}大過", _DisciplinaryMeasures.ToString()));
                    }
                }
                else
                {
                    if (歷年大過 >= _DisciplinaryMeasures)
                    {

                        check = true;
                        RemarkList.Add(string.Format("滿{0}大過", _DisciplinaryMeasures.ToString()));
                    }
                }
            }
            return check;
        }
    }

    class Crett
    {
        public int 歷年功過相抵A { get; set; }
        public int 歷年功過相抵B { get; set; }
        public int 歷年功過相抵C { get; set; }
    }
}
