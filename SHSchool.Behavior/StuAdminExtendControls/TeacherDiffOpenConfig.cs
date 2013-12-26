using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using Framework.Feature;
using System.Xml;
using K12.Data;
using FISCA.DSAUtil;
using DevComponents.DotNetBar.Controls;
using FISCA.LogAgent;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    public partial class TeacherDiffOpenConfig : BaseForm
    {
        private const string TimeDisplayFormat = "yyyy/MM/dd HH:mm";

        StringBuilder sb = new StringBuilder();

        public TeacherDiffOpenConfig()
        {
            InitializeComponent();

            XmlElement xml = Config.GetMoralUploadConfig();
            if (xml != null)
            {
                labelX1.Text += School.DefaultSchoolYear;
                labelX2.Text += School.DefaultSemester;
                //XmlNode n = xml.SelectSingleNode("SchoolYear");
                //if (n != null) integerInput1.Text = n.InnerText;

                //n = xml.SelectSingleNode("Semester");
                //if (n != null) integerInput2.Text = n.InnerText;

                XmlNode n = xml.SelectSingleNode("StartTime");
                if (n != null)
                {
                    DateTime? dt = DateTimeHelper.ParseGregorian(n.InnerText, PaddingMethod.First);
                    if (dt.HasValue)
                        textBoxX1.Text = dt.Value.ToString(TimeDisplayFormat);
                }

                n = xml.SelectSingleNode("EndTime");
                if (n != null)
                {
                    DateTime? dt = DateTimeHelper.ParseGregorian(n.InnerText, PaddingMethod.First);
                    if (dt.HasValue)
                        textBoxX2.Text = dt.Value.ToString(TimeDisplayFormat);
                }
            }

            sb.AppendLine("「開放時間設定」已被修改。");
            sb.AppendLine("修改前：");
            sb.AppendLine("開始時間「" + textBoxX1.Text + "」");
            sb.AppendLine("結束時間「" + textBoxX2.Text + "」");
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(errorProvider1.GetError(textBoxX1)) ||
                !string.IsNullOrEmpty(errorProvider2.GetError(textBoxX2)))
            {
                MsgBox.Show("資料請修正後再儲存!!");
                return;
            }

            if (DateTimeParsebifi())
            {
                string startTime = DateTime.Parse(textBoxX1.Text).ToString(DateTimeHelper.StdDateTimeFormat);
                string EndTime = DateTime.Parse(textBoxX2.Text).ToString(DateTimeHelper.StdDateTimeFormat);
                DSXmlHelper obj = new DSXmlHelper("Request");
                obj.AddElement("Content");
                obj.AddElement("Content", "StartTime", startTime);
                obj.AddElement("Content", "EndTime", EndTime);

                Config.SetMoralUploadConfig(obj.BaseElement);

                //LOG
                sb.AppendLine("修改後：");
                sb.AppendLine("開始時間「" + startTime + "」");
                sb.AppendLine("結束時間「" + EndTime + "」");
                ApplicationLog.Log("學務系統.開放時間設定", "修改開放時間設定", sb.ToString()); 

                Close();
            }
            else
            {
                MsgBox.Show("錯誤:\n開放時間大於截止時間!!");
                return;
            }
        }

        private bool DateTimeParsebifi()
        {
            if (DateTime.Parse(textBoxX1.Text).CompareTo(DateTime.Parse(textBoxX2.Text)) ==1)
            {
                return false;
            }
            else
            {
                return true;
            }
        }


        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBoxX1_Validating(object sender, CancelEventArgs e)
        {
            //是否空白
            if (textBoxX1.Text == string.Empty)
            {
                errorProvider1.SetError(textBoxX1, "請輸入日期!!");
                return;
            }
            else
            {
                errorProvider1.Clear();
            }

            CheckTimeMet(textBoxX1, errorProvider1);
        }

        private void textBoxX2_Validating(object sender, CancelEventArgs e)
        {
            if (textBoxX2.Text == string.Empty)
            {
                errorProvider2.SetError(textBoxX2, "請輸入日期!!");
                return;
            }
            else
            {
                errorProvider2.Clear();
            }

            CheckTimeMet(textBoxX2, errorProvider2);
        }

        private void CheckTimeMet(TextBoxX TBOX,ErrorProvider EP)
        {
            //是否為正確日期格式
            DateTime? dt = DateTimeHelper.ParseGregorian(TBOX.Text, PaddingMethod.First);
            if (dt.HasValue)
            {
                TBOX.Text = dt.Value.ToString(TimeDisplayFormat);
                EP.Clear();
            }
            else
            {
                EP.SetError(TBOX, "請輸入正確日期格式");
            }
        }

        private void textBoxX1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBoxX1_Validating(null, null);
            }
        }

        private void textBoxX2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBoxX2_Validating(null, null);
            }
        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxX1.Checked)
            {
                textBoxX1.Enabled = false;
                textBoxX1.Text = DateTime.Now.ToString(TimeDisplayFormat);
                textBoxX1_Validating(null, null);
            }
            else
            {
                textBoxX1.Enabled = true;
                textBoxX1.Text = "";
                textBoxX1_Validating(null, null);
            }
        }
    }
}
