using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.DSAUtil;
using Framework.Feature;
using System.Xml;
using SHSchool.Data;
using Campus.DocumentValidator;

namespace SHSchool.Behavior
{
    static public class GetConfig
    {

        //Period
        static public List<string> GetPeriodList()
        {
            List<string> list = new List<string>();

            foreach(K12.Data.PeriodMappingInfo info in K12.Data.PeriodMapping.SelectAll())
            {
                if (!list.Contains(info.Type))
                {
                    list.Add(info.Type);
                }
            }

            return list;
        }


        //Absence
        static public List<string> GetAbsenceList()
        {
            List<string> list = new List<string>();

            foreach (K12.Data.AbsenceMappingInfo info in K12.Data.AbsenceMapping.SelectAll())
            {
                if (!list.Contains(info.Name))
                {
                    list.Add(info.Name);
                }
            }

            return list;
        }


        /// <summary>
        /// 取得文字評量代碼表清單
        /// </summary>
        static public List<string> GetWordCommentList()
        {
            List<string> faceList = new List<string>();
            DSXmlHelper helper = Config.GetWordCommentList().GetContent();
            foreach (XmlElement morality in helper.GetElements("Content/Morality"))
            {
                DSXmlHelper moralityHelper = new DSXmlHelper(morality);
                string face = moralityHelper.GetText("@Face");
                if (!faceList.Contains(face))
                {
                    faceList.Add(face);
                }
            }
            return faceList;
        }

        /// <summary>
        /// 取得文字評量代碼表清單 and 詳細對照內容
        /// </summary>
        static public Dictionary<string, List<string>> GetWordCommentDic()
        {
            //List<string> faceList = new List<string>();
            Dictionary<string, List<string>> faceDic = new Dictionary<string, List<string>>();
            DSXmlHelper helper = Config.GetWordCommentList().GetContent();
            foreach (XmlElement morality in helper.GetElements("Content/Morality"))
            {
                DSXmlHelper moralityHelper = new DSXmlHelper(morality);
                string face = moralityHelper.GetText("@Face");
                //if (!faceList.Contains(face))
                //{
                //    faceList.Add(face);
                //}

                if (!faceDic.ContainsKey(face))
                {
                    faceDic.Add(face, new List<string>());
                }

                foreach (XmlElement item in moralityHelper.GetElements("Item"))
                {
                    DSXmlHelper dsxitem = new DSXmlHelper(item);
                    faceDic[face].Add(dsxitem.GetText("@Comment"));
                }
            }
            return faceDic;
        }

        static public Dictionary<string, string> GetMoralCommentCodeList()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            DSResponse dsrsp = Config.GetMoralCommentCodeList();
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Morality"))
            {
                if (!dic.ContainsKey(var.GetAttribute("Code")))
                {
                    dic.Add(var.GetAttribute("Code"), var.GetAttribute("Comment"));
                }
            }
            return dic;
        }
    }
}
