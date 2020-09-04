using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 日常生活表現預警表
{
    public partial class ViewMessage : BaseForm
    {
        public string _MessageName = "";
        public ViewMessage(string message)
        {
            InitializeComponent();
            _MessageName = message;
            textBoxX1.Text = _MessageName.Replace("<br>", "\r\n");
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            _MessageName = textBoxX1.Text.Replace("\r\n", "<br>");
            this.DialogResult = DialogResult.Yes;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            textBoxX1.Text = "您的電子報表已收到最新內容\r\n「{0}學年度 第{1}學期 日常生活表現預警表」";
        }
    }
}
