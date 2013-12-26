using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SHSchool.Data;

namespace SHSchool.Behavior
{
    class CommentLogRobot
    {
        public SHStudentRecord _student { get; set; }

        private bool Mode = false; //模式為true,就是新增資料

        private string _OldString { get; set; }

        private string _NewString { get; set; }

        private SHMoralScoreRecord _MSR { get; set; }

        public CommentLogRobot(SHStudentRecord student)
        {
            _student = student;
        }

        /// <summary>
        /// 註記此為新增資料
        /// </summary>
        public void SetOldLost()
        {
            Mode = true;
        }

        /// <summary>
        /// 修改前的XML資料
        /// </summary>
        public void SetOld(SHMoralScoreRecord MSR)
        {
            _OldString = MSR.Comment;
        }

        /// <summary>
        /// 修改後的XML資料
        /// </summary>
        public void SetNew(SHMoralScoreRecord MSR)
        {
            _MSR = MSR;
            _NewString = MSR.Comment;
        }

        /// <summary>
        /// 取得Log字串
        /// </summary>
        public string LogToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("學生「" + _student.Name + "」" + "學號「" + _student.StudentNumber + "」");
            if (Mode)
            {
                #region 新增導師評語
                sb.AppendLine("學年度「" + _MSR.SchoolYear + "」學期「" + _MSR.Semester + "」");
                sb.AppendLine("新增導師評語資料：");
                sb.AppendLine("「導師評語」已填入值「" + _NewString + "」");
                
                #endregion
                return sb.ToString();
            }
            else
            {
                #region 修改導師評語
                if (_OldString != _NewString)
                {                   
                    sb.AppendLine("學年度「" + _MSR.SchoolYear + "」學期「" + _MSR.Semester + "」");
                    sb.AppendLine("修改導師評語資料：");
                    sb.AppendLine("「導師評語」原有資料「" + _OldString + "」調整為「" + _NewString + "」");
                    return sb.ToString();
                }
                else
                {
                    return "";
                }
                #endregion
            }
        }
    }
}
