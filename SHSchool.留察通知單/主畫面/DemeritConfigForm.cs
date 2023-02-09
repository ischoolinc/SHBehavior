using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using DevComponents.DotNetBar.Controls;
using DevComponents.DotNetBar.Rendering;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;

namespace SHSchool.�d��q����
{
    public partial class DemeritConfigForm : BaseForm
    {
        private byte[] _buffer = null;
        private string base64 = null;
        private bool _isUpload = false;
        private string _defaultTemplate; //�w�]�d��1 + �w�]�d��2 + �ۭq�d��
        private DateRangeModeNew _mode = DateRangeModeNew.Month;
        private bool _printStudentList;

        public DemeritConfigForm(string defaultTemplate, DateRangeModeNew mode, byte[] buffer, string name, string address, bool printStudentList)
        {
            InitializeComponent();

            //�p�G�t�Ϊ�Renderer�OOffice2007Renderer
            //�P��_ClassTeacherView,_CategoryView���C��
            if (GlobalManager.Renderer is Office2007Renderer)
            {
                ((Office2007Renderer)GlobalManager.Renderer).ColorTableChanged += new EventHandler(ScoreCalcRuleEditor_ColorTableChanged);
                SetForeColor(this);
            }

            _defaultTemplate = defaultTemplate;
            _mode = mode;
            _printStudentList = printStudentList;

            if (buffer != null)
                _buffer = buffer;

            if (defaultTemplate == "�w�]�d��2") //�p�G�O2�~�]��2
                rbDEF_2.Checked = true;
            else if (defaultTemplate == "�ۭq�d��") //�p�G�O�ۭq
                radioButton2.Checked = true;
            else
                rbDEF_1.Checked = true; //�p�G�����O�N�i�J�w�]1

            checkBoxX2.Checked = printStudentList;

            switch (mode)
            {
                case DateRangeModeNew.Month:
                    radioButton3.Checked = true;
                    break;
                case DateRangeModeNew.Week:
                    radioButton4.Checked = true;
                    break;
                case DateRangeModeNew.Custom:
                    radioButton5.Checked = true;
                    break;
                default:
                    throw new Exception("Date Range Mode Error.");
            }

            //�]�w ComboBox
            Dictionary<ComboBoxEx, string> cboBoxes = new Dictionary<ComboBoxEx, string>();
            cboBoxes.Add(comboBoxEx1, name);
            cboBoxes.Add(comboBoxEx2, address);

            foreach (ComboBoxEx var in cboBoxes.Keys)
            {
                var.SelectedIndex = 0;
                foreach (DevComponents.Editors.ComboItem item in var.Items)
                {
                    if (item.Text == cboBoxes[var])
                    {
                        var.SelectedIndex = var.Items.IndexOf(item);
                        break;
                    }
                }
            }       
        }

        void ScoreCalcRuleEditor_ColorTableChanged(object sender, EventArgs e)
        {
            SetForeColor(this);
        }

        private void SetForeColor(Control parent)
        {
            foreach (Control var in parent.Controls)
            {
                if (var is RadioButton)
                    var.ForeColor = ((Office2007Renderer)GlobalManager.Renderer).ColorTable.CheckBoxItem.Default.Text;
                SetForeColor(var);
            }
        }

        private void rbDEF_1_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDEF_1.Checked)
            {
                //radioButton2.Checked = false;
                _defaultTemplate = "�w�]�d��1";
            }
        }

        private void rbDEF_2_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDEF_2.Checked)
            {
                //radioButton2.Checked = false;
                _defaultTemplate = "�w�]�d��2";
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                //radioButton1.Checked = false;
                _defaultTemplate = "�ۭq�d��";
            }
        }

        private void checkBoxX2_CheckedChanged(object sender, EventArgs e)
        {
            _printStudentList = checkBoxX2.Checked;
        }

        private void linkDef1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "�t�s�s��";
            sfd.FileName = "�d��q����_�d��1.doc";
            sfd.Filter = "Word�ɮ� (*.doc)|*.doc|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.�g�ٳq����_��}�����d��, 0, Properties.Resources.�g�ٳq����_��}�����d��.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkDef2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "�t�s�s��";
            sfd.FileName = "�d��q����_�d��2.doc";
            sfd.Filter = "Word�ɮ� (*.doc)|*.doc|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.�g�ٳq����, 0, Properties.Resources.�g�ٳq����.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkViewGeDin_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "�t�s�s��";
            sfd.FileName = "�ۭq�d��q����d��.doc";
            sfd.Filter = "Word�ɮ� (*.doc)|*.doc|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    if (Aspose.Words.Document.DetectFileFormat(new MemoryStream(_buffer)) == Aspose.Words.LoadFormat.Doc)
                        fs.Write(_buffer, 0, _buffer.Length);
                    else
                        fs.Write(Properties.Resources.�g�ٳq����_��}�����d��, 0, Properties.Resources.�g�ٳq����_��}�����d��.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkUpData_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "��ܦۭq���d��q����d��";
            ofd.Filter = "Word�ɮ� (*.doc)|*.doc";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (Aspose.Words.Document.DetectFileFormat(ofd.FileName) == Aspose.Words.LoadFormat.Doc)
                    {
                        FileStream fs = new FileStream(ofd.FileName, FileMode.Open);

                        byte[] tempBuffer = new byte[fs.Length];
                        fs.Read(tempBuffer, 0, tempBuffer.Length);
                        base64 = Convert.ToBase64String(tempBuffer);
                        _isUpload = true;
                        fs.Close();
                        MsgBox.Show("�W�Ǧ��\�C");
                    }
                    else
                        MsgBox.Show("�W���ɮ׮榡����");
                }
                catch
                {
                    MsgBox.Show("���w���|�L�k�s���C", "�}���ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region �x�s Preference

            ConfigData cd = K12.Data.School.Configuration["�d��q����_ForSH"];
            XmlElement config = cd.GetXml("XmlData", null);

            //XmlElement config = CurrentUser.Instance.Preference["�g�ٳq����"];

            if (config == null)
            {
                config = new XmlDocument().CreateElement("�d��q����");
            }

            config.SetAttribute("Default", _defaultTemplate);

            XmlElement customize = config.OwnerDocument.CreateElement("CustomizeTemplate");
            XmlElement mode = config.OwnerDocument.CreateElement("DateRangeMode");
            XmlElement receive = config.OwnerDocument.CreateElement("Receive");
            XmlElement conditions = config.OwnerDocument.CreateElement("Conditions");
            XmlElement PrintStudentList = config.OwnerDocument.CreateElement("PrintStudentList");

            PrintStudentList.SetAttribute("Checked", _printStudentList.ToString());
            config.ReplaceChild(PrintStudentList, config.SelectSingleNode("PrintStudentList"));

            if (_isUpload) //�p�G�O�ۭq�d��
            {
                customize.InnerText = base64;
                config.ReplaceChild(customize, config.SelectSingleNode("CustomizeTemplate"));
            }

            mode.InnerText = ((int)_mode).ToString();
            config.ReplaceChild(mode, config.SelectSingleNode("DateRangeMode"));

            receive.SetAttribute("Name", ((DevComponents.Editors.ComboItem)comboBoxEx1.SelectedItem).Text);
            receive.SetAttribute("Address", ((DevComponents.Editors.ComboItem)comboBoxEx2.SelectedItem).Text);
            if (config.SelectSingleNode("Receive") == null)
                config.AppendChild(receive);
            else
                config.ReplaceChild(receive, config.SelectSingleNode("Receive"));

            cd.SetXml("XmlData", config);
            cd.Save();

            #endregion

            this.DialogResult = DialogResult.OK;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                _mode = DateRangeModeNew.Month;
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                _mode = DateRangeModeNew.Week;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                _mode = DateRangeModeNew.Custom;
            }
        }
    }
}