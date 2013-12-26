using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;
using K12.Data;

namespace SHSchool.Behavior
{
    class StudentNumberRepeatValidator : IFieldValidator
    {
        Dictionary<string, List<StudentRecord>> StudentDic = new Dictionary<string, List<StudentRecord>>();
        public StudentNumberRepeatValidator()
        {
            foreach (StudentRecord student in Student.SelectAll())
            {
                if (student.Status == StudentRecord.StudentStatus.刪除)
                    continue;

                if (!StudentDic.ContainsKey(student.StudentNumber))
                {
                    StudentDic.Add(student.StudentNumber, new List<StudentRecord>());
                }
                StudentDic[student.StudentNumber].Add(student);
            }
        }

        #region IFieldValidator 成員

        public string Correct(string Value)
        {
            return Value;
        }

        public string ToString(string template)
        {
            return template;
        }

        public bool Validate(string Value)
        {
            if (StudentDic.ContainsKey(Value)) //包含此學號
            {
                if (StudentDic[Value].Count > 1) //多名學生為錯誤
                {
                    return false;
                }
            }
            return true;
        }

        #endregion
    }
}
