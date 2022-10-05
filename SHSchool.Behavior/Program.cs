using System;
using System.Collections.Generic;
using System.Text;
using FISCA.DSAUtil;
using System.Xml;
using System.Collections;
using FISCA.Presentation;
using FISCA.Permission;
using Framework;
using FISCA;
using SHSchool.Behavior.StuAdminExtendControls;
using SHSchool.Behavior.ClassExtendControls;
using SHSchool.Behavior.StudentExtendControls;
using SHSchool.Data;
using K12.Presentation;
using System.Windows.Forms;
using SHSchool.Behavior.StudentExtendControls.Reports.學生獎勵明細;
using Campus.DocumentValidator;

namespace SHSchool.Behavior
{
    public static class Program
    {
        [MainMethod("SHSchool.Behavior")]
        static public void Main()
        {
            //資料驗證(學號是否存在系統驗證)
            FactoryProvider.FieldFactory.Add(new FieldValidatorFactory());

            #region 其它

            RibbonBarItem StudentReports = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"];
            //StudentReports["報表"].Image = Properties.Resources.boolean_field_fav_64;
            StudentReports["報表"]["學務相關報表"]["學生獎勵明細"].Enable = Permissions.學生獎勵明細權限;
            StudentReports["報表"]["學務相關報表"]["學生獎勵明細"].Click += delegate
            {
                new SHSchool.Behavior.StudentExtendControls.Reports.學生獎勵明細.Report().Print();
            };

            StudentReports = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"];
            StudentReports["報表"]["學務相關報表"]["歷年功過及出席統計表"].Enable = Permissions.歷年功過及出席統計表權限;
            StudentReports["報表"]["學務相關報表"]["歷年功過及出席統計表"].Click += delegate
            {
                OverTheYearsStatisticsForm form = new OverTheYearsStatisticsForm();
                form.ShowDialog();
            };

            //匯出
            MenuButton rbItemExport = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["匯出"]["學務相關匯出"];
            rbItemExport["匯出文字評量"].Enable = Permissions.匯出文字評量權限;
            rbItemExport["匯出文字評量"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new ExportMoralScore();
                ExportMoralScoreUI wizard = new ExportMoralScoreUI(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            MenuButton rbItemImport = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["匯入"]["學務相關匯入"];
            rbItemImport["匯入文字評量"].Enable = Permissions.匯入文字評量權限;
            rbItemImport["匯入文字評量"].Click += delegate
            {
                ImportMoralScore wizard = new ImportMoralScore();
                wizard.Execute();
            };

            MenuButton rbItemExport1 = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["匯出"]["學務相關匯出"];
            rbItemExport["匯出導師評語"].Enable = Permissions.匯出導師評語權限;
            rbItemExport["匯出導師評語"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new ExportComment();
                ExportMoralScoreUI wizard = new ExportMoralScoreUI(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            MenuButton rbItemImport1 = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["匯入"]["學務相關匯入"];
            rbItemImport["匯入導師評語"].Enable = Permissions.匯入導師評語權限;
            rbItemImport["匯入導師評語"].Click += delegate
            {
                ImportComment wizard = new ImportComment();
                wizard.Execute();
            };

            #endregion

            RibbonBarItem rbItem = MotherForm.RibbonBarItems["學生", "學務"];

            rbItem["缺曠"].Image = Properties.Resources.desk_64;
            rbItem["缺曠"].Enable = (Permissions.缺曠權限 && NLDPanels.Student.SelectedSource.Count >= 1);
            rbItem["缺曠"].Click += delegate
            {
                if (1 == NLDPanels.Student.SelectedSource.Count)
                {
                    SingleEditor editor = new SingleEditor(SHStudent.SelectByID(K12.Presentation.NLDPanels.Student.SelectedSource[0]));
                    editor.ShowDialog();
                }
                else
                {
                    MutiEditor editor = new MutiEditor(K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource));
                    editor.ShowDialog();
                }
            };

            rbItem["長假登錄"].Image = Properties.Resources.desk_clock_64;
            rbItem["長假登錄"].Enable = (Permissions.長假登錄權限 && NLDPanels.Student.SelectedSource.Count >= 1);
            rbItem["長假登錄"].Click += delegate
            {
                if (1 <= NLDPanels.Student.SelectedSource.Count)
                {
                    TestSingleEditor SBStatistics = new TestSingleEditor(K12.Presentation.NLDPanels.Student.SelectedSource);
                    SBStatistics.ShowDialog();
                }
            };

            if (Permissions.缺曠權限)
            {
                K12.Presentation.NLDPanels.Student.ListPaneContexMenu["缺曠"].Image = Properties.Resources.desk_64;
                K12.Presentation.NLDPanels.Student.ListPaneContexMenu["缺曠"].Enable = (NLDPanels.Student.SelectedSource.Count >= 1);
                K12.Presentation.NLDPanels.Student.ListPaneContexMenu["缺曠"].Click += delegate
                {
                    if (1 == NLDPanels.Student.SelectedSource.Count)
                    {
                        SingleEditor editor = new SingleEditor(SHStudent.SelectByID(K12.Presentation.NLDPanels.Student.SelectedSource[0]));
                        editor.ShowDialog();
                    }
                    else
                    {
                        MutiEditor editor = new MutiEditor(K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource));
                        editor.ShowDialog();
                    }
                };
            }

            if (Permissions.長假登錄權限)
            {
                K12.Presentation.NLDPanels.Student.ListPaneContexMenu["長假登錄"].Image = Properties.Resources.desk_clock_64;
                K12.Presentation.NLDPanels.Student.ListPaneContexMenu["長假登錄"].Enable = (NLDPanels.Student.SelectedSource.Count >= 1);
                K12.Presentation.NLDPanels.Student.ListPaneContexMenu["長假登錄"].Click += delegate
                {
                    if (1 <= NLDPanels.Student.SelectedSource.Count)
                    {
                        TestSingleEditor SBStatistics = new TestSingleEditor(K12.Presentation.NLDPanels.Student.SelectedSource);
                        SBStatistics.ShowDialog();
                    }
                };
            }

            #region 事件

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                StudentReports["報表"]["學務相關報表"]["學生獎勵明細"].Enable = (Permissions.學生獎勵明細權限 && NLDPanels.Student.SelectedSource.Count >= 1);

                if (Permissions.缺曠權限)
                {
                    rbItem["缺曠"].Enable = (NLDPanels.Student.SelectedSource.Count >= 1);
                    K12.Presentation.NLDPanels.Student.ListPaneContexMenu["缺曠"].Enable = (NLDPanels.Student.SelectedSource.Count >= 1);
                }
                if (Permissions.長假登錄權限)
                {
                    rbItem["長假登錄"].Enable = (NLDPanels.Student.SelectedSource.Count >= 1);
                    K12.Presentation.NLDPanels.Student.ListPaneContexMenu["長假登錄"].Enable = (NLDPanels.Student.SelectedSource.Count >= 1);
                }
            };

            #endregion

            #region 學務作業

            RibbonBarItem Process = FISCA.Presentation.MotherForm.RibbonBarItems["學務作業", "資料統計"];
            //Process["其他"].Image = Properties.Resources.boolean_field_fav_64;
            Process["報表"].Image = Properties.Resources.paste_64;
            Process["報表"]["獎懲人數統計"].Enable = Permissions.公務統計報表權限;
            Process["報表"]["獎懲人數統計"].Click += delegate
            {
                new SHSchool.Behavior.StuAdminExtendControls.DisciplineStatistics();
            };

            RibbonBarItem TimeSeupt = FISCA.Presentation.MotherForm.RibbonBarItems["學務作業", "基本設定"];
            TimeSeupt["設定"].Image = Properties.Resources.sandglass_unlock_64;
            TimeSeupt["設定"].Size = RibbonBarButton.MenuButtonSize.Large;
            //Process["其他"].Image = Properties.Resources.boolean_field_fav_64;
            TimeSeupt["設定"]["開放時間設定"].Enable = Permissions.開放時間設定權限;
            TimeSeupt["設定"]["開放時間設定"].Click += delegate
            {
                new SHSchool.Behavior.StuAdminExtendControls.TeacherDiffOpenConfig().ShowDialog();
            };
            #endregion

            #region 報表

            //2022/3/11 - Dylan決定關閉此功能
            //RibbonBarItem ClassReports2 = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"];
            //ClassReports2["報表"]["學務相關報表"]["德行表現特殊學生名單"].Enable = Permissions.德行表現特殊學生名單權限;
            //ClassReports2["報表"]["學務相關報表"]["德行表現特殊學生名單"].Click += delegate
            //{
            //    DeXingStatistic statistic = new DeXingStatistic(K12.Presentation.NLDPanels.Class.SelectedSource.ToArray());
            //    statistic.ShowDialog();
            //};

            //RibbonBarItem StudentReports = K12.Presentation.NLDPanels.Student.RibbonBarItems["統計報表"];
            //StudentReports["報表"]["新功能"].Image = Properties.Resources.boolean_field_fav_64;
            //StudentReports["報表"]["新功能"]["懲戒通知單"].Enable = Permissions.學生懲戒通知單權限;
            //StudentReports["報表"]["新功能"]["懲戒通知單"].Click += delegate
            //{
            //    new DisciplineNotification(false);
            //};

            #endregion

            #region 註冊權限

            //Catalog ribbon = RoleAclSource.Instance["學生"]["報表"];
            //ribbon.Add(new RibbonFeature(Permissions.學生懲戒通知單, "懲戒通知單(new)"));

            Catalog ribbon = RoleAclSource.Instance["學生"]["功能按鈕"];
            ribbon.Add(new RibbonFeature(Permissions.缺曠, "缺曠"));
            ribbon.Add(new RibbonFeature(Permissions.長假登錄, "長假登錄"));
            ribbon.Add(new RibbonFeature(Permissions.匯出文字評量, "匯出文字評量"));
            ribbon.Add(new RibbonFeature(Permissions.匯入文字評量, "匯入文字評量"));

            ribbon.Add(new RibbonFeature(Permissions.匯出導師評語, "匯出導師評語"));
            ribbon.Add(new RibbonFeature(Permissions.匯入導師評語, "匯入導師評語"));

            Catalog ribbon1 = RoleAclSource.Instance["學生"]["報表"];
            ribbon1.Add(new RibbonFeature(Permissions.學生獎勵明細, "學生獎勵明細"));
            ribbon1.Add(new RibbonFeature(Permissions.歷年功過及出席統計表, "歷年功過及出席統計表"));

            //Catalog ribbon2 = RoleAclSource.Instance["班級"]["報表"];
            //ribbon2.Add(new RibbonFeature(Permissions.德行表現特殊學生名單, "德行表現特殊學生名單"));

            //ribbon = RoleAclSource.Instance["班級"]["報表"];
            //ribbon.Add(new RibbonFeature(Permissions.班級懲戒通知單, "懲戒通知單(new)"));

            ribbon = RoleAclSource.Instance["學務作業"]["功能按鈕"];
            ribbon.Add(new RibbonFeature(Permissions.公務統計報表, "公務統計報表"));
            ribbon.Add(new RibbonFeature(Permissions.開放時間設定, "開放時間設定"));

            #endregion
        }
    }
}
