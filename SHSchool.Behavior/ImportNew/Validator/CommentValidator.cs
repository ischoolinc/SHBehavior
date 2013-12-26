using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;
using System.Xml;

namespace SHSchool.Behavior
{
    class CommentValidator : IFieldValidator
    {
        #region IFieldValidator 成員
        Dictionary<string, string> CommentCodeDic = new Dictionary<string, string>();
        List<char> DotList = new List<char>();

        public CommentValidator(XmlElement validatorDescription)
        {
            DotList = GetDotList(validatorDescription);

            //取得導師評語代碼表
            CommentCodeDic = GetConfig.GetMoralCommentCodeList();
            //foreach (string each in dic.Keys)
            //{
            //    if (!dicValueKey.ContainsKey(dic[each]))
            //    {
            //        dicValueKey.Add(dic[each], each);
            //    }
            //}
        }

        /// <summary>
        /// 取得分割字典內容
        /// </summary>
        private List<char> GetDotList(XmlElement validatorDescription)
        {
            List<char> list = new List<char>();
            foreach (XmlElement each in validatorDescription.SelectNodes("Item"))
            {
                char[] now = each.GetAttribute("Value").ToCharArray();
                if (now.Count() == 1)
                {
                    if (!list.Contains((char)now.GetValue(0)))
                    {
                        list.Add((char)now.GetValue(0));
                    }
                }
            }
            return list;
        }

        //自動修正
        public string Correct(string Value)
        {
            //分割字串
            //string[] arry = Value.Split(',');
            //List<string> list = new List<string>();
            //foreach (string each in arry)
            //{
            //    if (dic.ContainsKey(each)) //如果是代碼則替換為字串
            //    {
            //        list.Add(dic[each]);
            //    }
            //    else
            //    {
            //        list.Add(each);
            //    }
            //}
            //return string.Join(",", list.ToArray());
            return Value;
        }

        public string ToString(string template)
        {
            return template;
        }

        public bool Validate(string Value)
        {
            Dictionary<char, bool> dic = new Dictionary<char, bool>();
                        
            foreach (char dot in DotList) //試圖掃瞄所有的分隔符號
            {
                bool check = true; //預設為True

                string[] arry = Value.Split(dot);
                foreach (string each in arry)
                {
                    //如果包含於Key或是Value之內(是代碼或是文字評量)
                    if (CommentCodeDic.ContainsValue(each))
                    {

                    }
                    else
                    {
                        check = false; //只要有一個錯就會警告
                    }
                }
                dic.Add(dot, check);
            }
            bool ReturnValue = false;
            foreach (char each in dic.Keys)
            {
                if (dic[each] == true)
                {
                    ReturnValue = dic[each];
                }
            }
            return ReturnValue;
        }

        #endregion
    }
}
