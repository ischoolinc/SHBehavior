using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SHSchool.Data;
using SmartSchool.API.PlugIn;
using FISCA.DSAUtil;
using System.Xml;

namespace SHSchool.Behavior
{

    class ExportMoralScore : SmartSchool.API.PlugIn.Export.Exporter
    {
        public ExportMoralScore()
        {
            this.Image = null;
            this.Text = "匯出文字評量";
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            //取得文字評量代碼表清單
            List<string> faceList = GetConfig.GetWordCommentList();
            //加入欲匯出之欄位資料
            wizard.ExportableFields.AddRange("學年度", "學期");
            foreach (string each in faceList)
            {
                wizard.ExportableFields.Add(each);
            }

            wizard.ExportPackage += delegate(object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                //取得學生清單
                List<SHStudentRecord> students = SHStudent.SelectByIDs(e.List);

                //取得學生相關文字評量記錄
                Dictionary<string, List<SHMoralScoreRecord>> DicMoralScore = GetMoralScore(e.List);

                students.Sort(SortStudent);

                //整理填入資料
                foreach (SHStudentRecord stud in students) //每一位學生
                {
                    if (DicMoralScore.ContainsKey(stud.ID))
                    {
                        foreach (SHMoralScoreRecord MScore in DicMoralScore[stud.ID]) //每一筆資料
                        {
                            //取得文字評量內容的Dictionary物件,可加快比對速度
                            Dictionary<string, string> studTextScore = GetTextScore(MScore);

                            RowData row = new RowData();
                            row.ID = stud.ID;
                            //對於每一個要匯出的欄位
                            foreach (string field in e.ExportFields)
                            {
                                if (wizard.ExportableFields.Contains(field))
                                {
                                    string value = "";

                                    switch(field) {
                                        case "學年度": 
                                            value= MScore.SchoolYear.ToString(); 
                                            break;
                                        case "學期": 
                                            value = MScore.Semester.ToString(); 
                                            break;
                                        default :
                                            value = findTextScore(studTextScore, field); //取得文字評量
                                            break;
                                    }
                                    row.Add(field, value); 
                                }
                            }
                            e.Items.Add(row);
                        }
                    }
                }
            };
        }

        private Dictionary<string, List<SHMoralScoreRecord>> GetMoralScore(List<string> list)
        {
            Dictionary<string, List<SHMoralScoreRecord>> DicMoralScore = new Dictionary<string, List<SHMoralScoreRecord>>();
            //依學生ID取得所有文字評量物件
            List<SHMoralScoreRecord> ListMoralScore = SHMoralScore.SelectByStudentIDs(list);

            foreach (SHMoralScoreRecord Mscore in ListMoralScore)
            {
                if (!DicMoralScore.ContainsKey(Mscore.RefStudentID))
                {
                    DicMoralScore.Add(Mscore.RefStudentID, new List<SHMoralScoreRecord>());
                }
                DicMoralScore[Mscore.RefStudentID].Add(Mscore);
            }

            return DicMoralScore;
        }

        //<Morality Face=\"誠實地面對自己\">橙</Morality>
        //<Morality Face=\"誠實地對待他人\">橙</Morality>
        //<Morality Face=\"自我尊重\">橙</Morality>
        //<Morality Face=\"尊重他人\">橙</Morality>
        //<Morality Face=\"承諾\">橙</Morality>
        //<Morality Face=\"責任\">橙</Morality>
        //<Morality Face=\"榮譽\">橙</Morality>
        //<Morality Face=\"感謝\">橙</Morality>
        private Dictionary<string, string> GetTextScore(SHMoralScoreRecord MScore)
        {
            Dictionary<string, string> studTextScore = new Dictionary<string, string>();
            foreach (XmlElement each in MScore.TextScore.SelectNodes("Morality"))
            {
                string key = each.Attributes["Face"].InnerText;
                if (!studTextScore.ContainsKey(key))
                {
                    studTextScore.Add(key, each.InnerText);
                }
            }
            return studTextScore;
        }

        private string findTextScore(Dictionary<string, string> studTextScore, string field)
        {
            string value = "";
            if (studTextScore.ContainsKey(field))
            {
                value = studTextScore[field];
            }            
            return value;
        }

        private int SortStudent(SHStudentRecord x, SHStudentRecord y)
        {

            string xx1 = x.Class != null ? x.Class.Name : "";
            string xx2 = x.SeatNo.HasValue ? x.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            string xx3 = xx1 + xx2;

            string yy1 = y.Class != null ? y.Class.Name : "";
            string yy2 = y.SeatNo.HasValue ? y.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            string yy3 = yy1 + yy2;

            return xx3.CompareTo(yy3);
        }
    }
}
