using FISCA;
using FISCA.Presentation;
using FISCA.Permission;

namespace 德行成績試算表
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            RibbonBarItem ClassReports = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"];
            ClassReports["報表"]["學務相關報表"]["日常表現記錄表(新制)"].Click += delegate
            {
                NewForm calc = new NewForm();
                calc.ShowDialog();
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                ClassReports["報表"]["學務相關報表"]["日常表現記錄表(新制)"].Enable = (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0) && Permissions.日常表現記錄表權限;
            };

            Catalog ribbon = RoleAclSource.Instance["班級"]["報表"];
            ribbon.Add(new RibbonFeature(Permissions.日常表現記錄表, "日常表現記錄表(新制)"));

        }
    }
}
