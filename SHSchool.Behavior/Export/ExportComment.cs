using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SHSchool.Data;
using SmartSchool.API.PlugIn;

namespace SHSchool.Behavior
{
    class ExportComment : SmartSchool.API.PlugIn.Export.Exporter
    {
        public ExportComment()
        {
            this.Image = null;
            this.Text = "匯出導師評語";
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange("學年度", "學期","導師評語");
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
                            RowData row = new RowData();
                            row.ID = stud.ID;
                            //對於每一個要匯出的欄位
                            foreach (string field in e.ExportFields)
                            {
                                if (wizard.ExportableFields.Contains(field))
                                {
                                    string value = "";

                                    switch (field)
                                    {
                                        case "學年度":
                                            value = MScore.SchoolYear.ToString();
                                            break;
                                        case "學期":
                                            value = MScore.Semester.ToString();
                                            break;
                                        default:
                                            value = MScore.Comment; //取得導師評語
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
