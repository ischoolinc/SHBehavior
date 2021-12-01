using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;
using FISCA.Presentation.Controls;

namespace SHSchool.留察通知單
{

    //SHSchool.KeptInSchoolAnAdviceNote.dll
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            #region 學生留察通知單
            
            RibbonBarItem StudentReports = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"];
            //StudentReports["報表"]["學務相關報表"].Image = Properties.Resources.boolean_field_fav_64;
            StudentReports["報表"]["學務相關報表"]["通知單"]["留察通知單"].Enable = Permissions.學生留察通知單權限;
            StudentReports["報表"]["學務相關報表"]["通知單"]["留察通知單"].Click += delegate
            {
                new SHSchool.留察通知單.Report("student").Print();
            };
            #endregion

            RibbonBarItem ClassReports = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"];
            //ClassReports["報表"]["學務相關報表"].Image = Properties.Resources.boolean_field_fav_64;
            ClassReports["報表"]["學務相關報表"]["通知單"]["留察通知單"].Enable = Permissions.班級留察通知單權限;
            ClassReports["報表"]["學務相關報表"]["通知單"]["留察通知單"].Click += delegate
            {
                new SHSchool.留察通知單.Report("class").Print();
            };

            #region 事件
            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count <= 0)
                {
                    StudentReports["報表"]["學務相關報表"]["通知單"]["留察通知單"].Enable = false;
                }
                else
                {
                    StudentReports["報表"]["學務相關報表"]["通知單"]["留察通知單"].Enable = Permissions.學生留察通知單權限;
                }
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count <= 0)
                {
                    ClassReports["報表"]["學務相關報表"]["通知單"]["留察通知單"].Enable = false;
                }
                else
                {
                    ClassReports["報表"]["學務相關報表"]["通知單"]["留察通知單"].Enable = Permissions.班級留察通知單權限;
                }
            };
            #endregion

            Catalog ribbon = RoleAclSource.Instance["學生"]["報表"];
            ribbon.Add(new RibbonFeature(Permissions.學生留察通知單, "留察通知單"));

            ribbon = RoleAclSource.Instance["班級"]["報表"];
            ribbon.Add(new RibbonFeature(Permissions.班級留察通知單, "留察通知單"));
        }
    }
}
