using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHSchool.DailyManifestation
{
    /// <summary>
    /// 代表目前使用者的相關權限資訊。
    /// </summary>
    public static class Permissions
    {

        public static string 學生個人表現總表 { get { return "SHSchool.DailyManifestation"; } }

        /// <summary>
        /// 班級缺曠記錄明細
        /// </summary>
        public static bool 學生個人表現總表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生個人表現總表].Executable;
            }
        }
    }
}
