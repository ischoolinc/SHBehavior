using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using K12.Data;
using K12.BusinessLogic;
using SHSchool.Data;
using System.Data;

namespace SHSchool.DailyManifestation
{
    class StudentInfo
    {
        public Dictionary<string, StudentDataObj> DicStudent = new Dictionary<string, StudentDataObj>();

        private List<StudentRecord> ListStudent = new List<StudentRecord>();
        private List<PhoneRecord> ListPhone = new List<PhoneRecord>();
        private List<ParentRecord> ListParent = new List<ParentRecord>();

        private List<AddressRecord> ListAddress = new List<AddressRecord>();
        private List<MeritRecord> ListMerit = new List<MeritRecord>();
        private List<DemeritRecord> ListDeMerit = new List<DemeritRecord>();

        private List<SHMoralScoreRecord> ListMoralScore = new List<SHMoralScoreRecord>();
        private List<UpdateRecordRecord> ListUpdataRecord = new List<UpdateRecordRecord>();
        private List<SemesterHistoryRecord> ListJHSemesterHistory = new List<SemesterHistoryRecord>();

        private List<AutoSummaryRecord> ListAutoSummary = new List<AutoSummaryRecord>();

        FISCA.Data.QueryHelper _QueryHelper = new FISCA.Data.QueryHelper();

        List<string> _BehaviorList;
        /// <summary>
        /// 學生資料整理器
        /// </summary>
        public StudentInfo(List<string> BehaviorList)
        {
            Dictionary<string, List<string>> merit_flagDic = SetMeritflag();
           
            _BehaviorList = BehaviorList;
            SetStudentBoxs();

            NewStudentData();

            SetPhone();
            SetParent();
            SetAddress();
            SetMeritList();
            SetDemeritList();
            SetMoralScoreList();
            SetUpdataList();
            SetSemesterHistory();
            SetAssnCode(); //社團
            SetAutoSummary(); //缺曠/獎懲統計

            foreach (string each in DicStudent.Keys)
            {
                DicStudent[each].SetupData(merit_flagDic);
            }

        }

        private Dictionary<string, List<string>> SetMeritflag()
        {
            Dictionary<string, List<string>> merit_flagDic = new Dictionary<string, List<string>>();
            //取得目前系統內,所有留察資料
            StringBuilder sb = new StringBuilder();
            sb.Append("select discipline.ref_student_id,discipline.school_year,discipline.semester from student join discipline ");
            sb.Append("on student.id = discipline.ref_student_id ");
            sb.Append("where discipline.merit_flag=2");
            DataTable dt = _QueryHelper.Select(sb.ToString());
            foreach (DataRow each in dt.Rows)
            {
                string studentID = "" + each[0];
                string schoolyear = "" + each[1];
                string semester = "" + each[2];
                if (!merit_flagDic.ContainsKey(studentID))
                {
                    merit_flagDic.Add(studentID, new List<string>());
                }
                merit_flagDic[studentID].Add(schoolyear + "/" + semester);
            }

            return merit_flagDic;
        }

        /// <summary>
        /// 取得所有資料
        /// </summary>
        private void SetStudentBoxs()
        {
            List<string> StudentIDList = K12.Presentation.NLDPanels.Student.SelectedSource;
            ListStudent = Student.SelectByIDs(StudentIDList); //取得學生
            ListStudent.Sort(new Comparison<StudentRecord>(ParseStudent));

            ListPhone = Phone.SelectByStudentIDs(StudentIDList); //取得電話資料
            ListParent = Parent.SelectByStudentIDs(StudentIDList); //取得監護人資料

            ListAddress = Address.SelectByStudentIDs(StudentIDList); //取得地址資料
            ListMerit = Merit.SelectByStudentIDs(StudentIDList); //取得獎勵資料(一對多)
            ListDeMerit = Demerit.SelectByStudentIDs(StudentIDList); //取得懲戒資料(一對多)
            ListMoralScore = SHMoralScore.SelectByStudentIDs(StudentIDList); //取得日常生活表現資料(一對多)
            ListUpdataRecord = UpdateRecord.SelectByStudentIDs(StudentIDList); //取得異動資料(一對多)
            ListJHSemesterHistory = SemesterHistory.SelectByStudentIDs(StudentIDList); //取得學生學期歷程
            //ListAssnCode = _accessHelper.Select<AssnCode>(); //取得所有社團記錄

            ListAutoSummary = AutoSummary.Select(StudentIDList, null);
            ListAutoSummary.Sort(SortSchoolYearSemester);
        }

        private int SortSchoolYearSemester(AutoSummaryRecord X, AutoSummaryRecord Y)
        {
            string SortX = X.SchoolYear.ToString() + X.Semester.ToString();
            string SortY = Y.SchoolYear.ToString() + Y.Semester.ToString();
            return SortX.CompareTo(SortY);
        }

        /// <summary>
        /// new一個學生資料模型
        /// </summary>
        private void NewStudentData()
        {
            foreach (StudentRecord stud in ListStudent)
            {
                if (!DicStudent.ContainsKey(stud.ID))
                {
                    DicStudent.Add(stud.ID, new StudentDataObj(_BehaviorList));
                    DicStudent[stud.ID].StudentRecord = stud;
                }
            }        
        }

        /// <summary>
        /// 填入電話資料
        /// </summary>
        private void SetPhone()
        {
            foreach (PhoneRecord phone in ListPhone)
            {
                if (DicStudent.ContainsKey(phone.RefStudentID))
                {
                    DicStudent[phone.RefStudentID].PhoneContact = phone.Contact;
                    DicStudent[phone.RefStudentID].PhonePermanent = phone.Permanent;
                }
            }
        }

        /// <summary>
        /// 取得監護人資料
        /// </summary>
        private void SetParent()
        {
            foreach (ParentRecord parent in ListParent)
            {
                if (DicStudent.ContainsKey(parent.RefStudentID))
                {
                    DicStudent[parent.RefStudentID].CustodianName = parent.Custodian.Name;
                    DicStudent[parent.RefStudentID].CustodianTitle = parent.Custodian.Relationship;
                }
            }
        }

        /// <summary>
        /// 填入地址資料
        /// </summary>
        private void SetAddress()
        {
            foreach (AddressRecord address in ListAddress)
            {
                if (DicStudent.ContainsKey(address.RefStudentID))
                {
                    DicStudent[address.RefStudentID].AddressMailing = JoinAddress(address.Mailing);
                    DicStudent[address.RefStudentID].AddressPermanent = JoinAddress(address.Permanent);
                }
            }
        }

        /// <summary>
        /// 取得地址資料
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        private string JoinAddress(K12.Data.AddressItem address)
        {
            return address.ZipCode + address.County + address.Town + address.District + address.Area + address.Detail;
        }

        /// <summary>
        /// 填入獎勵資料(一對多)
        /// </summary>
        private void SetMeritList()
        {
            foreach (MeritRecord merit in ListMerit)
            {
                if (DicStudent.ContainsKey(merit.RefStudentID))
                {
                    DicStudent[merit.RefStudentID].ListMerit.Add(merit);
                }
            }
        }

        /// <summary>
        /// 填入懲戒資料(一對多)
        /// </summary>
        private void SetDemeritList()
        {
            foreach (DemeritRecord demerit in ListDeMerit)
            {
                if (DicStudent.ContainsKey(demerit.RefStudentID))
                {
                    DicStudent[demerit.RefStudentID].ListDeMerit.Add(demerit);
                }
            }
        }

        /// <summary>
        /// 填入日常生活表現資料
        /// </summary>
        private void SetMoralScoreList()
        {
            foreach (SHMoralScoreRecord moralScore in ListMoralScore)
            {
                if (DicStudent.ContainsKey(moralScore.RefStudentID))
                {
                    DicStudent[moralScore.RefStudentID].ListMoralScore.Add(moralScore);
                }
            }
        }

        /// <summary>
        /// 填入異動記錄
        /// </summary>
        private void SetUpdataList()
        {
            foreach (UpdateRecordRecord UpdateRecord in ListUpdataRecord)
            {
                if (DicStudent.ContainsKey(UpdateRecord.StudentID))
                {
                    DicStudent[UpdateRecord.StudentID].ListUpdateRecord.Add(UpdateRecord);
                }
            }
        }

        /// <summary>
        /// 填入學期歷程
        /// </summary>
        private void SetSemesterHistory()
        {
            foreach (SemesterHistoryRecord semesterHistory in ListJHSemesterHistory)
            {
                if (DicStudent.ContainsKey(semesterHistory.RefStudentID))
                {
                    DicStudent[semesterHistory.RefStudentID].SemesterHistory = semesterHistory;
                }
            }
        }

        /// <summary>
        /// 填入社團活動資料
        /// </summary>
        private void SetAssnCode()
        {
            //foreach (AssnCode assn in ListAssnCode)
            //{
            //    if (DicStudent.ContainsKey(assn.StudentID))
            //    {
            //        DicStudent[assn.StudentID].ListAssnCode.Add(assn);
            //    }
            //}
        }

        /// <summary>
        /// 取得缺曠統計內容
        /// </summary>
        private void SetAutoSummary()
        {
            foreach (AutoSummaryRecord auto in ListAutoSummary)
            {
                if (DicStudent.ContainsKey(auto.RefStudentID))
                {
                    DicStudent[auto.RefStudentID].ListAutoSummary.Add(auto);
                }
            }
        }

        /// <summary>
        /// 排序功能
        /// </summary>
        static public int ParseStudent(StudentRecord x, StudentRecord y)
        {
            //取得班級名稱
            string Xstring = x.Class != null ? x.Class.Name : "";
            string Ystring = y.Class != null ? y.Class.Name : "";

            //取得座號
            int Xint = x.SeatNo.HasValue ? x.SeatNo.Value : 0;
            int Yint = y.SeatNo.HasValue ? y.SeatNo.Value : 0;
            //班級名稱加:號加座號(靠右對齊補0)
            Xstring += ":" + Xint.ToString().PadLeft(2, '0');
            Ystring += ":" + Yint.ToString().PadLeft(2, '0');

            return Xstring.CompareTo(Ystring);

        }
    }
}
