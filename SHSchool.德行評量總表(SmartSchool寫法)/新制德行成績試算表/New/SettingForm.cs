using DevComponents.DotNetBar.Controls;
using DevComponents.Editors.DateTimeAdv;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace 德行成績試算表
{
    public partial class SettingForm : BaseForm
    {
        private MemoryStream _template = null;
        private string base64 = null;
        private byte[] _buffer = null;
        string _ConfigName = "";

        private string _useDefaultTemplate = "範本1";
        ConfigData cd { get; set; }
        bool _isUpload { get; set; }
        public SettingForm(string configName, string useDefaultTemplate, byte[] buffer)
        {
            InitializeComponent();

            _ConfigName = configName;
            _useDefaultTemplate = useDefaultTemplate;

            if (buffer != null)
                _buffer = buffer;

            if (_useDefaultTemplate == "範本2") //如果是2才設為2
                rbDEF_2.Checked = true;
            else if (_useDefaultTemplate == "自訂範本") //如果是自訂
                radioButton2.Checked = true;

            LoadPreference();
        }

        private void LoadPreference()
        {
            #region 讀取 Preference

            cd = K12.Data.School.Configuration[_ConfigName];
            XmlElement config = cd.GetXml("XmlData", null);

            if (config != null)
            {
                _useDefaultTemplate = config.GetAttribute("Default");
                XmlElement customize = (XmlElement)config.SelectSingleNode("CustomizeTemplate");

                if (customize != null)
                {
                    string templateBase64 = customize.InnerText;
                    _buffer = Convert.FromBase64String(templateBase64);
                    _template = new MemoryStream(_buffer);
                }
            }
            else
            {
                #region 產生空白設定檔
                config = new XmlDocument().CreateElement("日常生活表現總表_班級");
                config.SetAttribute("Default", "範本1");
                XmlElement customize = config.OwnerDocument.CreateElement("CustomizeTemplate");
                config.AppendChild(customize);

                cd.SetXml("XmlData", config);

                _useDefaultTemplate = "範本1";

                #endregion
            }

            cd.Save(); //儲存組態資料。

            #endregion

        }

        private void linkDef1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "日常生活表現_班級(範本1).docx";
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.高中_日常生活表現_班級_a3, 0, Properties.Resources.高中_日常生活表現_班級_a3.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkDef2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "日常生活表現_班級(範本2).docx";
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.高中_日常生活表現_班級_版本2, 0, Properties.Resources.高中_日常生活表現_班級_版本2.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkViewGeDin_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "日常生活表現_班級(自訂範本).docx";
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Aspose.Words.Document doc = new Aspose.Words.Document(new MemoryStream(_buffer));
                    doc.Save(sfd.FileName, Aspose.Words.SaveFormat.Docx);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("檔案無法儲存。" + ex.Message);
                    return;
                }

                try
                {
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("檔案無法開啟。" + ex.Message);
                    return;
                }
            }



        }

        private void linkUpData_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            if (_buffer == null)
            {
                MsgBox.Show("目前沒有任何範本，請重新上傳。");
                return;
            }

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "選擇自訂的日常生活表現_班級(範本)";
            ofd.Filter = "Word檔案 (*.docx)|*.docx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    FileStream fs = new FileStream(ofd.FileName, FileMode.Open);

                    byte[] tempBuffer = new byte[fs.Length];
                    fs.Read(tempBuffer, 0, tempBuffer.Length);
                    base64 = Convert.ToBase64String(tempBuffer);
                    _isUpload = true;
                    fs.Close();
                    MsgBox.Show("上傳成功。\n(請按下儲存完成設定)");
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "開啟檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }



        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            #region 儲存 Preference

            ConfigData cd = K12.Data.School.Configuration[_ConfigName];
            XmlElement config = cd.GetXml("XmlData", null);

            //XmlElement config = CurrentUser.Instance.Preference["懲戒通知單"];

            if (config == null)
            {
                config = new XmlDocument().CreateElement("日常生活表現總表_班級");
            }

            config.SetAttribute("Default", _useDefaultTemplate);

            XmlElement customize = config.OwnerDocument.CreateElement("CustomizeTemplate");
          

            if (_isUpload) //如果是自訂範本
            {
                customize.InnerText = base64;
                config.ReplaceChild(customize, config.SelectSingleNode("CustomizeTemplate"));
            }

            cd.SetXml("XmlData", config);
            cd.Save();

            #endregion

            this.DialogResult = DialogResult.OK;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = string.Format("日常生活表現_班級_功能變數總表_{0}.docx", DateTime.Now.ToString("HHmmss"));
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.日常生活表現總表_功能變數總表, 0, Properties.Resources.日常生活表現總表_功能變數總表.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void rbDEF_1_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDEF_1.Checked)
            {
                //radioButton2.Checked = false;
                _useDefaultTemplate = "範本1";
            }
        }

        private void rbDEF_2_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDEF_2.Checked)
            {
                //radioButton2.Checked = false;
                _useDefaultTemplate = "範本2";
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                //radioButton1.Checked = false;
                _useDefaultTemplate = "自訂範本";
            }
        }
    }
}
