using System;
using System.Collections.Generic;
using System.Text;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.PlugIn.Report;
using SmartSchool.Customization.PlugIn;
using SmartSchool.Customization.Data.StudentExtension;

namespace SHSchool.Behavior.ClassExtendControls
{
    public class DisciplineStatistics
    {
        private List<string> mRewardCommentList;
        private List<RewardInfo> mRewardList;

        private string mMinReward;
        private int mMinRewardCount;
        
        private DateTime mStartDate;
        private DateTime mEndDate;

        private int mFaultACount = 0;
        private int mFaultBCount = 0;
        private int mFaultCCount = 0;

        private int mFaultASemesterCount = 0;
        private int mFaultBSemesterCount = 0;
        private int mFaultCSemesterCount = 0;


        private string GetRewardComment(RewardInfo Reward)
        {
            string strRewardComment = "";

            //統計所有獎懲記錄
            if (mMinReward.Equals(""))
            {
                //如果是懲戒記錄/如果未銷過
                if (Reward.Detail.SelectSingleNode("MeritFlag").InnerText == "0" && !Reward.Cleared)
                {
                    strRewardComment = Reward.OccurDate.Month.ToString() + "/" + Reward.OccurDate.Day.ToString();
                    strRewardComment += " " + Reward.OccurReason;

                    if (Reward.FaultA > 0)
                    {
                        strRewardComment += " 大過" + Reward.FaultA.ToString() + " 次";
                        mFaultACount += Reward.FaultA;
                    }

                    if (Reward.FaultB > 0)
                    {
                        //strRewardComment = Reward.OccurDate.Month.ToString() + "/" + Reward.OccurDate.Day.ToString();
                        //strRewardComment += " " + Reward.OccurReason;

                        strRewardComment += " 小過" + Reward.FaultB.ToString() + " 次";
                        mFaultBCount += Reward.FaultB;
                    }

                    if (Reward.FaultC > 0)
                    {
                        //strRewardComment = Reward.OccurDate.Month.ToString() + "/" + Reward.OccurDate.Day.ToString();
                        //strRewardComment += " " + Reward.OccurReason;

                        strRewardComment += " 警告" + Reward.FaultC.ToString() + " 次";
                        mFaultCCount += Reward.FaultC;
                    }
                }
                else if(Reward.Detail.SelectSingleNode("MeritFlag").InnerText == "2")//如果是留查資料
                {
                    strRewardComment = Reward.OccurDate.Month.ToString() + "/" + Reward.OccurDate.Day.ToString();
                    strRewardComment += " " + Reward.OccurReason +" (留校察看)";
                }
            }

            if (Reward.Detail.SelectSingleNode("MeritFlag").InnerText == "0" && !Reward.Cleared)
            {
                //警告
                if (mMinReward.Equals("警告"))
                {
                    strRewardComment = Reward.OccurDate.Month.ToString() + "/" + Reward.OccurDate.Day.ToString();
                    strRewardComment += " " + Reward.OccurReason;

                    if (Reward.FaultA > 0)
                    {
                        strRewardComment += " 大過" + Reward.FaultA.ToString() + " 次";
                        mFaultACount += Reward.FaultA;
                    }

                    if (Reward.FaultB > 0)
                    {
                        //strRewardComment = Reward.OccurDate.Month.ToString() + "/" + Reward.OccurDate.Day.ToString();
                        //strRewardComment += " " + Reward.OccurReason;

                        strRewardComment += " 小過" + Reward.FaultB.ToString() + " 次";
                        mFaultBCount += Reward.FaultB;
                    }

                    if (Reward.FaultC >= mMinRewardCount && Reward.FaultC > 0)
                    {
                        //strRewardComment = Reward.OccurDate.Month.ToString() + "/" + Reward.OccurDate.Day.ToString();
                        //strRewardComment += " " + Reward.OccurReason;

                        strRewardComment += " 警告" + Reward.FaultC.ToString() + " 次";
                        mFaultCCount += Reward.FaultC;
                    }
                }

                //小過
                if (mMinReward.Equals("小過"))
                {
                    strRewardComment = Reward.OccurDate.Month.ToString() + "/" + Reward.OccurDate.Day.ToString();
                    strRewardComment += " " + Reward.OccurReason;

                    if (Reward.FaultA > 0)
                    {
                        strRewardComment += " 大過" + Reward.FaultA.ToString() + " 次";
                        mFaultACount += Reward.FaultA;
                    }

                    if (Reward.FaultB >= mMinRewardCount && Reward.FaultB > 0)
                    {
                        //strRewardComment = Reward.OccurDate.Month.ToString() + "/" + Reward.OccurDate.Day.ToString();
                        //strRewardComment += " " + Reward.OccurReason;

                        strRewardComment += " 小過" + Reward.FaultB.ToString() + " 次";
                        mFaultBCount += Reward.FaultB;
                    }
                }

                //大過
                if (mMinReward.Equals("大過"))
                {
                    strRewardComment = Reward.OccurDate.Month.ToString() + "/" + Reward.OccurDate.Day.ToString();
                    strRewardComment += " " + Reward.OccurReason;


                    if (Reward.FaultA >= mMinRewardCount && Reward.FaultA > 0)
                    {

                        strRewardComment += " 大過" + Reward.FaultA.ToString() + " 次";
                        mFaultACount += Reward.FaultA;
                    }
                }
            }

            return strRewardComment;
        }

        private void SetSemesterCount(RewardInfo Reward)
        {
            //如果是懲戒記錄/如果未銷過
            if (Reward.Detail.SelectSingleNode("MeritFlag").InnerText == "0" && !Reward.Cleared)
            {
                //統計所有獎懲
                if (mMinReward.Equals(""))
                {
                    if (Reward.FaultA > 0)
                        mFaultASemesterCount += Reward.FaultA;

                    if (Reward.FaultB > 0)
                        mFaultBSemesterCount += Reward.FaultB;

                    if (Reward.FaultC > 0)
                        mFaultCSemesterCount += Reward.FaultC;
                }

                //警告
                if (mMinReward.Equals("警告"))
                {
                    if (Reward.FaultA > 0)
                        mFaultASemesterCount += Reward.FaultA;

                    if (Reward.FaultB > 0)
                        mFaultBSemesterCount += Reward.FaultB;

                    if (Reward.FaultC >= mMinRewardCount && Reward.FaultC > 0)
                        mFaultCSemesterCount += Reward.FaultC;
                }

                //小過
                if (mMinReward.Equals("小過"))
                {
                    if (Reward.FaultA > 0)
                        mFaultASemesterCount += Reward.FaultA;

                    if (Reward.FaultB >= mMinRewardCount && Reward.FaultB > 0)
                        mFaultBSemesterCount += Reward.FaultB;
                }

                //大過
                if (mMinReward.Equals("大過"))
                {
                    if (Reward.FaultA >= mMinRewardCount && Reward.FaultA > 0)
                        mFaultASemesterCount += Reward.FaultA;
                }
            } 
        }

        public DisciplineStatistics(List<RewardInfo> RewardList, string MinReward, int MinRewardCount, string StartDate, string EndDate)
        {
            mRewardList = RewardList;
            mMinReward = MinReward;
            mMinRewardCount = MinRewardCount;

            if (!DateTime.TryParse(StartDate,out mStartDate))
                mStartDate=DateTime.Now;

            if (!DateTime.TryParse(EndDate,out mEndDate))
                mEndDate=DateTime.Now;

            mRewardCommentList = new List<string>();

            //獎懲記錄
            foreach (RewardInfo Reward in mRewardList)
            {
                //過濾掉獎勵記錄
                if (Reward.Detail.SelectSingleNode("MeritFlag").InnerText != "1")
                {
                    if (Reward.SchoolYear == SmartSchool.Customization.Data.SystemInformation.SchoolYear && Reward.Semester == SmartSchool.Customization.Data.SystemInformation.Semester)
                        SetSemesterCount(Reward);

                    if (Reward.OccurDate >= mStartDate && Reward.OccurDate <= mEndDate)
                    {
                        string RewardComment = GetRewardComment(Reward);
                        if (!RewardComment.Equals(""))
                            mRewardCommentList.Add(RewardComment);
                    }
                }
            }

        }

        //取得獎懲列表
        public List<string> RewardCommentList
        {
            get
            {
                return mRewardCommentList;
            }
        }

        //取得大過數量，依據使用者選取期間來統計
        public int FaultACount
        {
            get
            {
                return mFaultACount;
            }
        }

        //取得小過數量，依據使用者選取期間來統計
        public int FaultBCount
        {
            get
            {
                return mFaultBCount;
            }
        }

        //取得警告數量，依據使用者選取期間來統計
        public int FaultCCount
        {
            get
            {
                return mFaultCCount;
            }
        }

        //取得大過數量，依據本學期來統計
        public int FaultASemesterCount
        {
            get
            {
                return mFaultASemesterCount;
            }
        }

        //取得小過數量，依據本學期來統計
        public int FaultBSemesterCount
        {
            get
            {
                return mFaultBSemesterCount;
            }
        }

        //取得警告數量，依據本學期來統計
        public int FaultCSemesterCount
        {
            get
            {
                return mFaultCSemesterCount;
            }
        }
    }
}