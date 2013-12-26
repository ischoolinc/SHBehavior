using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SHSchool.Data;
using System.Xml;

namespace SHSchool.Behavior
{
    class LogRobot
    {
        public SHStudentRecord _student { get; set; }
        
        private List<string> _List { get; set; }
        //舊資料
        private XmlElement _OldInfo { get; set; }
        //新資料
        private XmlElement _NewInfo { get; set; }

        private Dictionary<string, string> _OldDic { get; set; }

        private Dictionary<string, string> _NewDic { get; set; }

        private SHMoralScoreRecord _MSR { get; set; }

        private bool Mode = false; //模式為true,就是新增資料

        /// <summary>
        /// 建構子
        /// </summary>
        public LogRobot(SHStudentRecord student, List<string> list)
        {
            _student = student;
            _List = list;
            _OldDic = new Dictionary<string, string>();
            _NewDic = new Dictionary<string, string>();
            foreach (string each in _List)
            {
                if (!_OldDic.ContainsKey(each))
                {
                    _OldDic.Add(each, "");
                }
                if (!_NewDic.ContainsKey(each))
                {
                    _NewDic.Add(each, "");
                }
            }
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
            _MSR = MSR;
            _OldInfo = _MSR.TextScore;

            if (_OldInfo != null)
            {
                foreach (XmlElement each in _OldInfo.SelectNodes("Morality"))
                {
                    if (_List.Contains(each.GetAttribute("Face")))
                    {
                        _OldDic[each.GetAttribute("Face")] = each.InnerText;
                    }
                }
            }
        }

        /// <summary>
        /// 修改後的XML資料
        /// </summary>
        public void SetNew(SHMoralScoreRecord MSR)
        {
            _MSR = MSR;
            _NewInfo = _MSR.TextScore;
            if (_NewInfo != null)
            {
                foreach (XmlElement each in _NewInfo.SelectNodes("Morality"))
                {
                    if (_List.Contains(each.GetAttribute("Face")))
                    {
                        _NewDic[each.GetAttribute("Face")] = each.InnerText;
                    }
                }
            }
        }

        /// <summary>
        /// 取得Log字串
        /// </summary>
        /// <returns></returns>
        public string LogToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("學生「" + _student.Name + "」" + "學號「" + _student.StudentNumber + "」");
            if (Mode)
            {
                #region 新增文字評量資料
                sb.AppendLine("學年度「" + _MSR.SchoolYear + "」學期「" + _MSR.Semester + "」");
                sb.AppendLine("新增文字評量資料：");
                foreach (string each in _List)
                {
                    sb.AppendLine("欄位「" + each + "」已填入值「" + _NewDic[each] + "」");
                }   
                #endregion         
            }
            else
            {
                #region 修改文字評量資料
                sb.AppendLine("學年度「" + _MSR.SchoolYear + "」學期「" + _MSR.Semester + "」");
                sb.AppendLine("修改文字評量資料：");

                bool UpdateChenge = false;
                foreach (string each in _List)
                {
                    if (_OldDic[each] != _NewDic[each])
                    {
                        UpdateChenge = true;
                    }
                }

                if (UpdateChenge)
                {
                    foreach (string each in _List)
                    {
                        if (_OldDic[each] != _NewDic[each]) //如果不一樣則修改資料
                        {
                            sb.AppendLine("欄位「" + each + "」原有資料「" + _OldDic[each] + "」調整為「" + _NewDic[each] + "」");
                        }
                    }
                }
                else
                {
                    sb.AppendLine("資料未有任何修改!!");
                } 
                #endregion
            } 
            return sb.ToString();
        }
    }        
}
