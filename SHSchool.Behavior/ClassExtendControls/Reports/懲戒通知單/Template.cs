using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

namespace SHSchool.Behavior.ClassExtendControls
{
    public class DisciplineNotificationTemplate : ISchoolDocument
    {

        private int mStartRowIndex = 0;
        private string mRawTemplatePath;
        private Aspose.Words.Document mtemplate;

        public DisciplineNotificationTemplate(string RawTemplatePath)
        {
            mRawTemplatePath = RawTemplatePath;
            mtemplate = new Aspose.Words.Document(mRawTemplatePath, LoadFormat.Doc, "");
        }

        public DisciplineNotificationTemplate(System.IO.Stream Stream)
        {
            mtemplate = new Aspose.Words.Document(Stream);
        }

        public int ProcessDocument()
        {

            return 0;
        }

        public object ExtraInfo(string value)
        {
            if (value.Equals("StartRowIndex"))
                return mStartRowIndex;
            else
                return null;
        }

        public Aspose.Words.Document Document
        {
            get
            {
                return mtemplate;
            }
        }
    }
}