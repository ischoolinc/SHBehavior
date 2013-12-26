using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Permission;

namespace SHSchool.DailyManifestation
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            string toolName = "學生日常表現總表";

            K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["學務相關報表"][toolName].Enable = false;
            K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["學務相關報表"][toolName].Click += delegate
            {
                NewSRoutineForm sr = new NewSRoutineForm();
                sr.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                bool SelectNow = (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.學生個人表現總表權限);
                K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["學務相關報表"][toolName].Enable = SelectNow;
            };

            Catalog detail = RoleAclSource.Instance["學生"]["報表"];
            detail.Add(new RibbonFeature(Permissions.學生個人表現總表, toolName));

        }



    }
}
