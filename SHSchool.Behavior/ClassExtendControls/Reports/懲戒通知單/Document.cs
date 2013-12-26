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


namespace SHSchool.Behavior.ClassExtendControls
{
    public class DisciplineNotificationDocument : ISchoolDocument
    {
        private StudentRecord mstudent;
        private DisciplineNotificationConfig mconfig;
        private ISchoolDocument mtemplate;
        private Aspose.Words.Document mdoc;

        public DisciplineNotificationDocument(StudentRecord student, DisciplineNotificationConfig config, DisciplineNotificationTemplate template)
        {
            mstudent = student;
            mconfig = config;
            mtemplate = template;
        }

        public int ProcessDocument()
        {
            Document eachSection = new Document();
            eachSection.Sections.Clear();
            eachSection.Sections.Add(eachSection.ImportNode(mtemplate.Document.Sections[0], true));

            //基本資料
            DisciplineNotificationRecord ANR = new DisciplineNotificationRecord(mstudent, mconfig);
            //獎懲狀態
            DisciplineStatistics RS = new DisciplineStatistics(mstudent.RewardList, mconfig.MinReward, mconfig.MinRewardCount, mconfig.StartDate, mconfig.EndDate);

            string[] key = new string[] { "1", "2", "3", "4", "5", "收件人地址", "收件人姓名", "學校名稱", "學校地址", "學校電話", "資料期間", "班級", "座號", "學號", "學生姓名", "導師", "學期累計大過", "學期累計小過", "學期累計警告","本期累計大過", "本期累計小過", "本期累計警告", "懲戒明細" };
            object[] value = new object[] { ANR.ZipCode01, ANR.ZipCode02, ANR.ZipCode03, "", "", ANR.ReceiverAddress, ANR.Receiver, ANR.SchoolName, ANR.SchoolAddress, ANR.SchoolTel, ANR.DataRange, ANR.ClassName, ANR.SeatNo, ANR.StudentNumber, ANR.StudentName, ANR.TeacherName, RS.FaultASemesterCount, RS.FaultBSemesterCount, RS.FaultCSemesterCount, RS.FaultACount, RS.FaultBCount, RS.FaultCCount, RS.RewardCommentList};

            //合併列印
            eachSection.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(AbsenceNotification_MailMerge_MergeField);
            eachSection.MailMerge.RemoveEmptyParagraphs = true;
            eachSection.MailMerge.Execute(key, value);
            mdoc = eachSection;

            return RS.FaultACount + RS.FaultBCount + RS.FaultCCount;
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


        private void AbsenceNotification_MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            if (e.FieldName == "懲戒明細")
            {
                List<string> eachStudentDisciplineDetail = (List<string>)e.FieldValue;

                Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(e.Document);

                builder.MoveToField(e.Field, false);
                builder.StartTable();
                builder.CellFormat.ClearFormatting();
                builder.CellFormat.Borders.ClearFormatting();
                builder.CellFormat.VerticalAlignment = Aspose.Words.CellVerticalAlignment.Center;
                builder.CellFormat.LeftPadding = 3.0;
                builder.RowFormat.LeftIndent = 0.0;
                builder.RowFormat.Height = 15.0;

                int rowNumber = 6;

                if (eachStudentDisciplineDetail.Count > rowNumber * 2)
                {
                    rowNumber += (eachStudentDisciplineDetail.Count - (rowNumber * 2)) / 2;
                    rowNumber += (eachStudentDisciplineDetail.Count - (rowNumber * 2)) % 2;
                }

                for (int j = 0; j < rowNumber; j++)
                {
                    builder.InsertCell();
                    builder.CellFormat.Borders.Right.LineStyle = Aspose.Words.LineStyle.Single;
                    builder.CellFormat.Borders.Right.Color = Color.Black;
                    if (j < eachStudentDisciplineDetail.Count)
                        builder.Write(eachStudentDisciplineDetail[j]);
                    builder.InsertCell();
                    if (j + rowNumber < eachStudentDisciplineDetail.Count)
                        builder.Write(eachStudentDisciplineDetail[j + rowNumber]);
                    builder.EndRow();
                }

                builder.EndTable();

                e.Text = string.Empty;
            }
        }

    }
}