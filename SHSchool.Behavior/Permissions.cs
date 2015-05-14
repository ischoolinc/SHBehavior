using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Framework;

namespace SHSchool.Behavior
{
    /// <summary>
    /// 代表目前使用者的相關權限資訊。
    /// </summary>
    public static class Permissions
    {

        //public static string 學生懲戒通知單 { get { return "SHSchool.Behavior.Student.DisciplineNotificationForm"; } }
        //public static string 班級懲戒通知單 { get { return "SHSchool.Behavior.Class.DisciplineNotificationForm"; } }
        public static string 公務統計報表 { get { return "SHSchool.Behavior.StuAdmin.DisciplineStatistics"; } }
        public static string 開放時間設定 { get { return "SHSchool.Behavior.StuAdmin.TeacherDiffOpenConfig"; } }

        public static string 學生獎勵明細 { get { return "SHSchool.Behavior.Student.SelectMeritForm"; } }

        public static string 缺曠 { get { return "K12.Student.AttendanceForm"; } }
        public static string 長假登錄 { get { return "K12.Student.TestSingleEditor5"; } }

        public static string 德行表現特殊學生名單 { get { return "Report0190"; } }

        public static string 匯出文字評量 { get { return "SHSchool.Behavior.Student.Export.MoralScore"; } }
        public static string 匯入文字評量 { get { return "SHSchool.Behavior.Student.Import.MoralScore"; } }

        public static string 匯出導師評語 { get { return "SHSchool.Behavior.Student.Export.Comment"; } }
        public static string 匯入導師評語 { get { return "SHSchool.Behavior.Student.Import.Comment"; } }

        public static bool 匯入導師評語權限
        {
            get { return FISCA.Permission.UserAcl.Current[匯入導師評語].Executable; }
        }

        public static bool 匯出導師評語權限
        {
            get { return FISCA.Permission.UserAcl.Current[匯出導師評語].Executable; }
        }

        public static bool 匯入文字評量權限
        {
            get { return FISCA.Permission.UserAcl.Current[匯入文字評量].Executable; }
        }

        public static bool 匯出文字評量權限
        {
            get { return FISCA.Permission.UserAcl.Current[匯出文字評量].Executable; }
        }

        public static bool 德行表現特殊學生名單權限
        {
            get { return FISCA.Permission.UserAcl.Current[德行表現特殊學生名單].Executable; }
        }

        public static bool 缺曠權限
        {
            get { return FISCA.Permission.UserAcl.Current[缺曠].Executable; }
        }

        public static bool 長假登錄權限
        {
            get { return FISCA.Permission.UserAcl.Current[長假登錄].Executable; }
        }

        public static bool 學生獎勵明細權限
        {
            get { return FISCA.Permission.UserAcl.Current[學生獎勵明細].Executable; }
        }
 
        //public static bool 學生懲戒通知單權限 { get { return FISCA.Permission.UserAcl.Current[學生懲戒通知單].Executable; } }
        //public static bool 班級懲戒通知單權限 { get { return FISCA.Permission.UserAcl.Current[班級懲戒通知單].Executable; } }


        public static bool 公務統計報表權限 { get { return FISCA.Permission.UserAcl.Current[公務統計報表].Executable; } }


        public static bool 開放時間設定權限 { get { return FISCA.Permission.UserAcl.Current[開放時間設定].Executable; } }


        public static string 歷年功過及出席統計表 { get { return "Report0030"; } }
        public static bool 歷年功過及出席統計表權限
        {
            get { return FISCA.Permission.UserAcl.Current[歷年功過及出席統計表].Executable; }
        }
    }
}
