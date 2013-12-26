using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;

namespace SHSchool.Behavior.ClassExtendControls
{

    public enum DateRangeMode { Month, Week, Custom };

    public class DisciplineNotificationPreference
    {
        static private DisciplineNotificationPreference mDisciplineNotificationPreference = null;

        private System.Xml.XmlElement mconfig = null;
        private MemoryStream mDefaultTemplate = new MemoryStream(Properties.Resources.獎懲通知單_住址位移);

        //產生空白設定檔
        public void InitialPreference()
        {
            mconfig = new XmlDocument().CreateElement("懲戒通知單");
            mconfig.SetAttribute("Default", "true");

            XmlElement printSetup = mconfig.OwnerDocument.CreateElement("PrintHasRecordOnly");
            XmlElement customize = mconfig.OwnerDocument.CreateElement("CustomizeTemplate");
            XmlElement dateRangeMode = mconfig.OwnerDocument.CreateElement("DateRangeMode");
            XmlElement receive = mconfig.OwnerDocument.CreateElement("Receive");

            printSetup.SetAttribute("Checked", "false");
            dateRangeMode.InnerText = ((int)DateRangeMode.Month).ToString();

            receive.SetAttribute("Name", "");
            receive.SetAttribute("Address", "");

            mconfig.AppendChild(printSetup);
            mconfig.AppendChild(customize);
            mconfig.AppendChild(dateRangeMode);
            mconfig.AppendChild(receive);

            SmartSchool.Customization.Data.SystemInformation.Preference["懲戒通知單"] = mconfig;
        }

        public void Refresh()
        {
            //從SmartSchool當中取得設定變數，若是為null則重新設定

            mconfig = SmartSchool.Customization.Data.SystemInformation.Preference["懲戒通知單"];

            if (mconfig == null)
                InitialPreference();
        }

        private DisciplineNotificationPreference()
        {
            Refresh();
        }

        static public DisciplineNotificationPreference GetInstance()
        {
            if (mDisciplineNotificationPreference == null)
                mDisciplineNotificationPreference = new DisciplineNotificationPreference();


            return mDisciplineNotificationPreference;
        }

        public bool UseAdvanceCondition
        {
            get
            {
                if (mconfig.GetAttribute("AdvanceCondition").Equals(""))
                    mconfig.SetAttribute("AdvanceCondition", "true");

                return bool.Parse(mconfig.GetAttribute("AdvanceCondition"));
            }
            set
            {
                mconfig.SetAttribute("AdvanceCondition", value.ToString());
            }
        }

        public bool UseDefaultTemplate
        {
            get
            {
                if (mconfig.GetAttribute("Default").Equals(""))
                    mconfig.SetAttribute("Default", "true");

                return bool.Parse(mconfig.GetAttribute("Default"));
            }
            set
            {
                mconfig.SetAttribute("Default", value.ToString());
            }
        }

        public string CustomizeTemplateString
        {
            set
            {
                XmlElement CustomizeTemplateElm = mconfig.OwnerDocument.CreateElement("CustomizeTemplate");
                CustomizeTemplateElm.InnerText = value;
                mconfig.ReplaceChild(CustomizeTemplateElm, mconfig.SelectSingleNode("CustomizeTemplate"));
            }
        }

        public byte[] CustomizeTemplateByte
        {
            get
            {
                byte[] Buffer = null;
                XmlElement customize = (XmlElement)mconfig.SelectSingleNode("CustomizeTemplate");

                if (customize != null)
                {
                    string templateBase64 = customize.InnerText;
                    Buffer = Convert.FromBase64String(templateBase64);
                    return Buffer;
                }
                else
                    return null;
            }
        }

        public MemoryStream CustomizeTemplate
        {
            get
            {
                MemoryStream Template = null;
                byte[] Buffer = null;
                XmlElement customize = (XmlElement)mconfig.SelectSingleNode("CustomizeTemplate");

                if (customize != null)
                {
                    string templateBase64 = customize.InnerText;
                    Buffer = Convert.FromBase64String(templateBase64);
                    Template = new MemoryStream(Buffer);

                    return Template;
                }
                else
                    return null;
            }
        }

        public bool PrintHasRecordOnly
        {
            get
            {
                XmlElement PrintElm = (XmlElement)mconfig.SelectSingleNode("PrintHasRecordOnly");

                if (PrintElm != null)
                    if (PrintElm.HasAttribute("Checked"))
                    {
                        string strChecked = PrintElm.GetAttribute("Checked");

                        if (strChecked.Equals(""))
                            PrintElm.SetAttribute("Checked", "false");

                        return bool.Parse(PrintElm.GetAttribute("Checked"));
                    }
                return true;
            }
            set
            {
                XmlElement PrintElm = mconfig.OwnerDocument.CreateElement("PrintHasRecordOnly");
                PrintElm.SetAttribute("Checked", value.ToString());

                if (mconfig.SelectSingleNode("PrintHasRecordOnly") != null)
                    mconfig.ReplaceChild(PrintElm, mconfig.SelectSingleNode("PrintHasRecordOnly"));
                else
                    mconfig.AppendChild(PrintElm);

                SmartSchool.Customization.Data.SystemInformation.Preference["懲戒通知單"] = mconfig;
            }
        }

        private string InitalReceiveData()
        {
            XmlElement newReceive = mconfig.OwnerDocument.CreateElement("Receive");
            newReceive.SetAttribute("Name", "");
            newReceive.SetAttribute("Address", "");
            mconfig.AppendChild(newReceive);
            SmartSchool.Customization.Data.SystemInformation.Preference["懲戒通知單"] = mconfig;

            return "";
        }

        public string ReceiveName
        {
            get
            {
                XmlElement ReceiveElm = (XmlElement)mconfig.SelectSingleNode("Receive");

                return (ReceiveElm != null) ? ReceiveElm.GetAttribute("Name") : InitalReceiveData();
            }
            set
            {
                if ((XmlElement)mconfig.SelectSingleNode("Receive") == null)
                    InitalReceiveData();

                XmlElement ReceiveElm = (XmlElement)mconfig.SelectSingleNode("Receive");

                ReceiveElm.SetAttribute("Name", value);
            }
        }

        public string ReceiveAddress
        {
            get
            {
                XmlElement ReceiveElm = (XmlElement)mconfig.SelectSingleNode("Receive");

                return (ReceiveElm != null) ? ReceiveElm.GetAttribute("Address") : InitalReceiveData();
            }
            set
            {
                if ((XmlElement)mconfig.SelectSingleNode("Receive") == null)
                    InitalReceiveData();

                XmlElement ReceiveElm = (XmlElement)mconfig.SelectSingleNode("Receive");

                ReceiveElm.SetAttribute("Address", value);
            }
        }


        private DateRangeMode InitialDateRangeMode()
        {
            XmlElement newDateRangeMode = mconfig.OwnerDocument.CreateElement("DateRangeMode");
            newDateRangeMode.InnerText = ((int)DateRangeMode.Month).ToString();
            mconfig.AppendChild(newDateRangeMode);
            SmartSchool.Customization.Data.SystemInformation.Preference["懲戒通知單"] = mconfig;

            return DateRangeMode.Month;
        }

        public DateRangeMode DateModeRangeMode
        {
            get
            {
                XmlElement DateRangeModeElm = (XmlElement)mconfig.SelectSingleNode("DateRangeMode");

                return (DateRangeModeElm != null) ? (DateRangeMode)int.Parse(DateRangeModeElm.InnerText) : InitialDateRangeMode();
            }
            set
            {
                XmlElement DateRangeModeElm = mconfig.OwnerDocument.CreateElement("DateRangeMode");
                DateRangeModeElm.InnerText = ((int)value).ToString();
                mconfig.ReplaceChild(DateRangeModeElm, mconfig.SelectSingleNode("DateRangeMode"));

            }
        }

        public string MinReward
        {
            get
            {
                return mconfig.GetAttribute("MinReward");
            }
            set
            {
                mconfig.SetAttribute("MinReward", value);
            } 
        }

        public int Mincount
        {
            get
            {
                int vMinCount = 0;

                int.TryParse(mconfig.GetAttribute("MinCount"),out vMinCount);

                return vMinCount;
            }
            set
            {
                mconfig.SetAttribute("MinCount", value.ToString());
            }
        }

        public string StartDate
        {
            get
            {
                return mconfig.GetAttribute("StartDate");
            }
            set
            {
                mconfig.SetAttribute("StartDate", value);
            }
        }

        public string EndDate
        {
            get
            {
                return mconfig.GetAttribute("EndDate");
            }
            set
            {
                mconfig.SetAttribute("EndDate", value);
            }
        }
    }
}