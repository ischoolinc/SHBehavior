using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHSchool.留察通知單
{
    class Permissions
    {
        public static string 學生留察通知單 { get { return "SHSchool.Behavior.Student.KeptInSchoolAnAdviceNote"; } }
        public static string 班級留察通知單 { get { return "SHSchool.Behavior.Class.KeptInSchoolAnAdviceNote"; } }

        public static bool 學生留察通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生留察通知單].Executable;
            }
        }

        public static bool 班級留察通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級留察通知單].Executable;
            }
        }

    }
}
