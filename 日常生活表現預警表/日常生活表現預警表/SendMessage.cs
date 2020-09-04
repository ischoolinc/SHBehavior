using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace 日常生活表現預警表
{
    class SendMessage
    {
        public enum Type { Student, Teacher }

        List<string> _tagIDList { get; set; }

        string _message { get; set; }

        string _msgTitle { get; set; }

        private Type _type { get; set; }

        BackgroundWorker bgw = new BackgroundWorker();

        /// <summary>
        /// 傳送推播通知
        /// </summary>
        public SendMessage(Type type, List<string> tagIDList, string message)
        {
            _type = type;
            _tagIDList = tagIDList;
            _message = message;

            bgw.DoWork += Bgw_DoWork;
            bgw.RunWorkerCompleted += Bgw_RunWorkerCompleted;

        }

        /// <summary>
        /// 開始傳送
        /// </summary>
        public void Run()
        {
            bgw.RunWorkerAsync();
        }

        private void Bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            BatchPushNotice();
        }

        private void BatchPushNotice()
        {
            //必須要使用greening帳號登入才能用
            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            XmlElement root = doc.CreateElement("Request");

            //標題
            var eleTitle = doc.CreateElement("Title");
            eleTitle.InnerText = "日常生活預警表已送達";
            root.AppendChild(eleTitle);

            //發送人員
            var eleDisplaySender = doc.CreateElement("DisplaySender");
            eleDisplaySender.InnerText = "電子報表通知";
            root.AppendChild(eleDisplaySender);

            //背景
            if (_message != "")
            {

                var eleMessage = doc.CreateElement("Message");
                eleMessage.InnerText = _message;
                root.AppendChild(eleMessage);

                foreach (string each in _tagIDList)
                {
                    //發送對象
                    var eleTarget = doc.CreateElement("TargetTeacher");
                    eleTarget.InnerText = each;
                    root.AppendChild(eleTarget);
                }

                //送出
                FISCA.DSAClient.XmlHelper xmlHelper = new FISCA.DSAClient.XmlHelper(root);
                var conn = new FISCA.DSAClient.Connection();
                conn.Connect(FISCA.Authentication.DSAServices.AccessPoint, "1campus.notice.admin.v17", FISCA.Authentication.DSAServices.PassportToken);
                var resp = conn.SendRequest("PostNotice", xmlHelper);


                //處理Log
                StringBuilder sb_log = new StringBuilder();
                List<string> byAction = new List<string>();

                byAction.Add("老師");

                sb_log.AppendLine(string.Format("發送對象「{0}」", string.Join("，", byAction)));
                sb_log.AppendLine(string.Format("發送標題「{0}」", "日常生活預警表已傳送"));
                sb_log.AppendLine(string.Format("發送單位「{0}」", "電子報表通知"));
                sb_log.AppendLine(string.Format("發送內容「{0}」", _message));
                sb_log.AppendLine(string.Format("對象清單「{0}」", string.Join(",", _tagIDList)));

                FISCA.LogAgent.ApplicationLog.Log("智慧簡訊", "發送", sb_log.ToString());


            }
            else
            {
                MsgBox.Show("請輸入內容");
            }
        }

        private void Bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                MsgBox.Show("發送成功!");
            }
            else
            {
                //前景
                MsgBox.Show("發送失敗:\n" + e.Error.Message);
            }
        }

    }
}
