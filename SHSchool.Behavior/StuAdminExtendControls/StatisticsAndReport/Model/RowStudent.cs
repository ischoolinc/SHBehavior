using System.Windows.Forms;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    internal class RowStudent : DataGridViewRow
    {
        private Student _student;
        public Student Student
        {
            get { return _student; }
        }
    }
}