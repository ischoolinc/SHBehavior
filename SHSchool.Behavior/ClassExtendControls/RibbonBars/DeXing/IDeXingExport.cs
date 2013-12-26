using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace SHSchool.Behavior.ClassExtendControls
{
    interface IDeXingExport
    {
        Control MainControl { get;}
        void LoadData();
        void Export();
    }
}
