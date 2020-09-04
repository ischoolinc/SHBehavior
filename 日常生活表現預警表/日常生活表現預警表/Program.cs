using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 日常生活表現預警表
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            RibbonBarItem Print = FISCA.Presentation.MotherForm.RibbonBarItems["班級", "資料統計"];
            Print["報表"]["學務相關報表"]["日常生活表現預警表"].Enable = Permissions.日常生活表現預警表權限;
            Print["報表"]["學務相關報表"]["日常生活表現預警表"].Click += delegate
            {
                PerformanceAlertForm paf = new PerformanceAlertForm();
                paf.ShowDialog();
            };

            Catalog ribbon = RoleAclSource.Instance["班級"]["報表"];
            ribbon.Add(new RibbonFeature(Permissions.日常生活表現預警表, "日常生活表現預警表"));
        }
    }
}
