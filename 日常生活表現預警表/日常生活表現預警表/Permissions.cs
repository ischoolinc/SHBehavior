using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 日常生活表現預警表
{
    class Permissions
    {
        public static string 日常生活表現預警表 { get { return "Daily_LifePerformanceAlertForm.cs"; } }
        public static bool 日常生活表現預警表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[日常生活表現預警表].Executable;
            }
        }
    }
}
