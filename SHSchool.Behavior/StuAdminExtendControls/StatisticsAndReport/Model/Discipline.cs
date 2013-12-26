using System.Collections.Generic;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    internal class Discipline
    {
        public Discipline(string student_id, string kind)
        {
            _student_id = student_id;
            _kind = kind;
        }

        private string _student_id;
        public string StudentID
        {
            get { return _student_id; }
        }

        private string _kind;
        public string Kind
        {
            get { return _kind; }
        }

        private bool _cleared = false;
        public bool Cleared
        {
            get { return _cleared; }
            set { _cleared = value; }
        }

        private int _awardA = 0;
        public int AwardA
        {
            get { return _awardA; }
            set { _awardA = value; }
        }

        private int _awardB = 0;
        public int AwardB
        {
            get { return _awardB; }
            set { _awardB = value; }
        }

        private int _awardC = 0;
        public int AwardC
        {
            get { return _awardC; }
            set { _awardC = value; }
        }

        private int _faultA = 0;
        public int FaultA
        {
            get { return _faultA; }
            set { _faultA = value; }
        }

        private int _faultB = 0;
        public int FaultB
        {
            get { return _faultB; }
            set { _faultB = value; }
        }

        private int _faultC = 0;
        public int FaultC
        {
            get { return _faultC; }
            set { _faultC = value; }
        }
    }

    internal class DisciplineCollection : List<Discipline>
    {
        
    }
}
