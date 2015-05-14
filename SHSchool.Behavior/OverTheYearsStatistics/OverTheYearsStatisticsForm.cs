using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using FISCA.DSAUtil;
using SmartSchool.Common;
using K12.Data.Configuration;

namespace SHSchool.Behavior
{
    public partial class OverTheYearsStatisticsForm : BaseForm
    {
        private bool _print_cleared = false;
        private Dictionary<string, List<string>> _print_types = new Dictionary<string, List<string>>();
        //設定1
        private string _ConfigValue2 = "歷年功過及出席統計";

        private string _configDefName = "節次分類";

        public OverTheYearsStatisticsForm()
        {
            InitializeComponent();
            LoadPreference();
        }

        private void LoadPreference()
        {
            #region 讀取 Preference
            XmlElement config = SmartSchool.Customization.Data.SystemInformation.Preference[_ConfigValue2];
            if (config != null)
            {
                //列印資訊
                XmlElement print = (XmlElement)config.SelectSingleNode("Print");
                if (print != null)
                {
                    DSXmlHelper printHelper = new DSXmlHelper(print);
                    _print_cleared = bool.Parse(printHelper.GetText("@IsPrintCleared"));
                    checkBoxX1.Checked = _print_cleared;
                }
                else
                {
                    XmlElement newPrint = config.OwnerDocument.CreateElement("Print");
                    newPrint.SetAttribute("IsPrintCleared", "False");
                    _print_cleared = false;
                    config.AppendChild(newPrint);
                    SmartSchool.Customization.Data.SystemInformation.Preference[_ConfigValue2] = config;
                }
            }
            else
            {
                #region 產生空白設定檔
                config = new XmlDocument().CreateElement(_ConfigValue2);
                SmartSchool.Customization.Data.SystemInformation.Preference[_ConfigValue2] = config;
                #endregion
            }

            //使用者設定的假別
            _print_types.Clear();

            K12.Data.Configuration.ConfigData Defconfig = K12.Data.School.Configuration[_ConfigValue2];

             XmlElement config_2;

             if (Defconfig.GetXml(_configDefName, null) == null)
             {
                 config_2 = new XmlDocument().CreateElement(_configDefName);
                 Defconfig[_configDefName] = config.OuterXml;
                 Defconfig.Save();
             }
             else
             {
                 config_2 = Defconfig.GetXml(_configDefName, null);

                 foreach (XmlElement type in config_2.SelectNodes("Type"))
                 {
                     string typeName = type.GetAttribute("Text");

                     if (!_print_types.ContainsKey(typeName))
                         _print_types.Add(typeName, new List<string>());

                     foreach (XmlElement absence in type.SelectNodes("Absence"))
                     {
                         string absenceName = absence.GetAttribute("Text");

                         if (!_print_types[typeName].Contains(absenceName))
                             _print_types[typeName].Add(absenceName);
                     }
                 }

             }




            #endregion
        }

        private void SavePreference()
        {
            #region 儲存 Preference

            XmlElement config = SmartSchool.Customization.Data.SystemInformation.Preference[_ConfigValue2];

            if (config == null)
                config = new XmlDocument().CreateElement(_ConfigValue2);

            XmlElement print = config.OwnerDocument.CreateElement("Print");
            print.SetAttribute("IsPrintCleared", checkBoxX1.Checked.ToString());

            if (config.SelectSingleNode("Print") == null)
                config.AppendChild(print);
            else
                config.ReplaceChild(print, config.SelectSingleNode("Print"));

            SmartSchool.Customization.Data.SystemInformation.Preference[_ConfigValue2] = config;

            #endregion
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            SavePreference();
            LoadPreference();
            new OverTheYearsStatistics(_print_cleared, _print_types);
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SelectTypeForm form = new SelectTypeForm(_ConfigValue2);
            if (form.ShowDialog() == DialogResult.OK)
            {
                LoadPreference();
            }
        }
    }
}