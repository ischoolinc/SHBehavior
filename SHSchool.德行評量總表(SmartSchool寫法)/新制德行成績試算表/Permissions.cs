using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 德行成績試算表
{
    class Permissions
    {
        public static string 日常表現記錄表 { get { return "SHSchool.SemesterMoralScoreCalcForm"; } }

        public static bool 日常表現記錄表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[日常表現記錄表].Executable;
            }
        }
    }
}
