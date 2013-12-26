using System.Collections.Generic;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    internal class Student
    {
        public Student(string id)
        {
            _id = id;
        }

        private string _id;
        public string ID
        {
            get { return _id; }
        }

        private string _name = "";
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        private string _student_number = "";
        public string StudentNumber
        {
            get { return _student_number; }
            set { _student_number = value; }
        }

        private string _gender = "";
        public string Gender
        {
            get { return _gender; }
            set { _gender = value; }
        }

        private string _grade_year = "";
        public string GradeYear
        {
            get { return _grade_year; }
            set { _grade_year = value; }
        }

        private string _dept = "";
        public string Department
        {
            get { return _dept; }
            set { _dept = value; }
        }
            
        private string _class = "";
        public string ClassName
        {
            get { return _class; }
            set { _class = value; }
        }

        private string _location = "";
        public string Location
        {
            get { return _location; }
            set { _location = value; }
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
            get { return _faultC;}
            set { _faultC = value; }
        }

        private bool _ultimateAdmonition = false;
        public bool UltimateAdmonition
        {
            get { return _ultimateAdmonition; }
            set { _ultimateAdmonition = value; }
        }
    }

    internal class StudentCollection : Dictionary<string, Student>
    {

    }
}
