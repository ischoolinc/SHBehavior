using System;
using System.Collections.Generic;
using System.Text;
using SHSchool.Behavior.StuAdminExtendControls;
using SHSchool.Behavior.Feature;

namespace SHSchool.Behavior.StudentExtendControls
{
    public class PeriodCollection
    {
        public PeriodCollection()
        {
            _periodList = new List<PeriodInfo>();
        }

        private List<PeriodInfo> _periodList;

        public List<PeriodInfo> Items
        {
            get { return _periodList; }
        }

        public List<PeriodInfo> GetSortedList()
        {
            _periodList.Sort(Compare);
            return _periodList;
        }

        public int Compare(PeriodInfo x, PeriodInfo y)
        {
            if (x.Sort == y.Sort) return 0;
            else if (x.Sort > y.Sort) return 1;
            else return -1;
        }
    }
}
