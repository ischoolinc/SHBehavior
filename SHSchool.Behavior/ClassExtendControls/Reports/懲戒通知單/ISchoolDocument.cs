using System;
using System.Collections.Generic;
using System.Text;

namespace SHSchool.Behavior.ClassExtendControls
{
    interface ISchoolDocument
    {

        int ProcessDocument();

        object ExtraInfo(string value);

        Aspose.Words.Document Document
        {
            get;
        }

    }
}