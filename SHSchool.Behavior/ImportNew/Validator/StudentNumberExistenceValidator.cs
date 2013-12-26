using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;
using K12.Data;

namespace SHSchool.Behavior
{
    class StudentNumberExistenceValidator : IFieldValidator
    {
        Dictionary<string, List<StudentRecord>> StudentDic = new Dictionary<string, List<StudentRecord>>();
        public StudentNumberExistenceValidator()
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

        //自動修正
        public string Correct(string Value)
        {
            return Value;
        }
        //回傳驗證訊息
        public string ToString(string template)
        {
            return template;
        }

        public bool Validate(string Value)
        {
            if (StudentDic.ContainsKey(Value)) //包含此學號
            {
                return true;
            }
            else //不包含此學號
            {
                return false;
            }
        }

        #endregion
    }
}
