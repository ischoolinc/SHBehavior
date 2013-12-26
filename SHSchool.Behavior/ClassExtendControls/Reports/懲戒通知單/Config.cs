using System;
using System.Collections.Generic;
using System.Text;

namespace SHSchool.Behavior.ClassExtendControls
{
    public class DisciplineNotificationConfig
    {
        private string mStartDate;
        private string mEndDate;
        private string mReceiverType;
        private string mReceiverAddressType;
        private DisciplineNotificationPreference mPreference;
        private string mMinReward;
        private int mMinRewardCount;

        public DisciplineNotificationConfig()
        {
            mStartDate = DateTime.Now.ToShortDateString();
            mEndDate = DateTime.Now.ToShortDateString();
            mReceiverType = "學生姓名";
            mReceiverAddressType = "戶籍地址";
            mMinReward = "警告";
            mMinRewardCount = 0;
            
        }

        public DisciplineNotificationConfig(DisciplineNotificationPreference Preference)
        {
            mMinReward = Preference.MinReward;
            mMinRewardCount = Preference.Mincount;
            mReceiverType = Preference.ReceiveName;
            mReceiverAddressType = Preference.ReceiveAddress;
            mStartDate = Preference.StartDate;
            mEndDate = Preference.EndDate;
            mPreference = Preference;
        }

        public DisciplineNotificationPreference Preference
        {
            get 
            {
                return mPreference;
            }
        }

        public string MinReward
        {
            get
            {
                return mMinReward;
            }
        }

        public int MinRewardCount
        {
            get
            {
                return mMinRewardCount;
            }
        }

        public string StartDate
        {
            get
            {
                return mStartDate;
            }
            set
            {
                mStartDate = value;
            }
        }

        public string EndDate
        {
            get
            {
                return mEndDate;
            }
            set
            {
                mEndDate = value;
            }
        }

        public string ReceiverType
        {
            get
            {
                return mReceiverType;
            }
            set
            {
                mReceiverType = value;
            }
        }

        public string ReceiverAddressType
        {
            get
            {
                return mReceiverAddressType;
            }
            set
            {
                mReceiverAddressType = value;
            }
        }
    }
}