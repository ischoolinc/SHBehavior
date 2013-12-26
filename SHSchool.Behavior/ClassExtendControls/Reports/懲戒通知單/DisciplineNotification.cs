using System;
using System.Collections.Generic;
using System.Text;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.PlugIn.Report;
using SmartSchool.Customization.PlugIn;
using SmartSchool.Customization.Data.StudentExtension;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace SHSchool.Behavior.ClassExtendControls
{
    public class DisciplineNotification : ISchoolDocument
    {
        //private ButtonAdapter classButton;
        //private ButtonAdapter studentButton;
        private Aspose.Words.Document mdoc;
        private DisciplineNotificationConfig mconfig;
        private DisciplineNotificationTemplate template;
        private DisciplineNotificationPreference mPreference;
        private bool IsClass;

        public DisciplineNotification(bool IsStudentOrClass)
        {
            IsClass = IsStudentOrClass;
            ProcessButtonClick();
        }

        public int ProcessDocument()
        {
            AccessHelper helper = new AccessHelper(); //helper物件用來取得ischool相關資料。

            mdoc = new Document(); //產生新的文件，用來合併每位學生的懲戒通知單
            mdoc.Sections.Clear(); //先將新文件的節做清除

            //判斷是否要使用預設的樣版或是自訂樣版
            if (mPreference.UseDefaultTemplate)
                template = new DisciplineNotificationTemplate(new MemoryStream(Properties.Resources.獎懲通知單_住址位移));
            else
                template = new DisciplineNotificationTemplate(mPreference.CustomizeTemplate);

            template.ProcessDocument();

            List<StudentRecord> students = new List<StudentRecord>();

            if (IsClass)
                foreach (ClassRecord crecord in helper.ClassHelper.GetSelectedClass())
                    students.AddRange(crecord.Students);
            else
                students.AddRange(helper.StudentHelper.GetSelectedStudent());

            //填入獎懲資訊
            helper.StudentHelper.FillReward(students);

            //填入連絡資訊
            helper.StudentHelper.FillContactInfo(students);

            //填入家長資訊
            helper.StudentHelper.FillParentInfo(students);

            //循訪每位學生記錄，並建立ActivityNotificationDocument物件來產生報表
            foreach (StudentRecord student in students)
            {
                ISchoolDocument studentdoc = new DisciplineNotificationDocument(student, mconfig, template);

                //將單一學生的每月生活通知單加入到主要的報表當中
                if (studentdoc.ProcessDocument() > 0)
                {
                    mdoc.Sections.Add(mdoc.ImportNode(studentdoc.Document.Sections[0], true));
                }
            }

            return 0;
        }

        public object ExtraInfo(string value)
        {
            return null;
        }

        public Aspose.Words.Document Document
        {
            get
            {
                return mdoc;
            }
        }

        public void ProcessButtonClick()
        {
            DemeritNotificatForm DateRangeForm = new DemeritNotificatForm();

            if (DateRangeForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                mPreference = DisciplineNotificationPreference.GetInstance();
                mconfig = new DisciplineNotificationConfig(mPreference);
                ProcessDocument();
                try
                {
                    #region 儲存並開啟檔案

                    string reportName = "懲戒通知單";
                    string path = Path.Combine(Application.StartupPath, "Reports");
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    path = Path.Combine(path, reportName + ".doc");

                    if (File.Exists(path))
                    {
                        int i = 1;
                        while (true)
                        {
                            string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                            if (!File.Exists(newPath))
                            {
                                path = newPath;
                                break;
                            }
                        }
                    }

                    try
                    {
                        Document.Save(path, SaveFormat.Doc);
                        System.Diagnostics.Process.Start(path);
                    }
                    catch
                    {
                        SaveFileDialog sd = new SaveFileDialog();
                        sd.Title = "另存新檔";
                        sd.FileName = reportName + ".doc";
                        sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                        if (sd.ShowDialog() == DialogResult.OK)
                        {
                            try
                            {
                                Document.Save(sd.FileName, SaveFormat.AsposePdf);
                            }
                            catch
                            {

                                MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }
                    #endregion
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            } 
        }
    }
}