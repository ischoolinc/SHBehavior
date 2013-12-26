using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHSchool.Behavior
{
    public static class SortClassIndex
    {

        #region SHClassRecord / SHStudentRecord

        static public List<SHSchool.Data.SHClassRecord> JHSchoolData_JHClassRecord(List<SHSchool.Data.SHClassRecord> ClassList)
        {
            ClassList.Sort(SortSHSchoolData_SHClassRecord);
            return ClassList;
        }

        static private int SortSHSchoolData_SHClassRecord(SHSchool.Data.SHClassRecord class1, SHSchool.Data.SHClassRecord class2)
        {
            string ClassYear1 = class1.GradeYear.HasValue ? class1.GradeYear.Value.ToString().PadLeft(10, '0') : string.Empty.PadLeft(10, '9');
            string ClassYear2 = class2.GradeYear.HasValue ? class2.GradeYear.Value.ToString().PadLeft(10, '0') : string.Empty.PadLeft(10, '9');

            string DisplayOrder1 = "";
            if (string.IsNullOrEmpty(class1.DisplayOrder))
            {
                DisplayOrder1 = class1.DisplayOrder.PadLeft(10, '9');
            }
            else
            {
                DisplayOrder1 = class1.DisplayOrder.PadLeft(10, '0');
            }
            string DisplayOrder2 = "";
            if (string.IsNullOrEmpty(class2.DisplayOrder))
            {
                DisplayOrder2 = class2.DisplayOrder.PadLeft(10, '9');
            }
            else
            {
                DisplayOrder2 = class2.DisplayOrder.PadLeft(10, '0');
            }

            string ClassName1 = class1.Name.PadLeft(10, '0');
            string ClassName2 = class2.Name.PadLeft(10, '0');

            string Compareto1 = ClassYear1 + DisplayOrder1 + ClassName1;
            string Compareto2 = ClassYear2 + DisplayOrder2 + ClassName2;

            return Compareto1.CompareTo(Compareto2);
        }

        static public List<SHSchool.Data.SHStudentRecord> SHSchoolData_SHStudentRecord(List<SHSchool.Data.SHStudentRecord> StudentList)
        {
            //整理出學生&班級資料清單
            List<string> classIDList = new List<string>();
            foreach (SHSchool.Data.SHStudentRecord student in StudentList)
            {
                if (!string.IsNullOrEmpty(student.RefClassID))
                {
                    if (!classIDList.Contains(student.RefClassID))
                    {
                        classIDList.Add(student.RefClassID);
                    }
                }
            }
            //一次取得班級清單
            List<SHSchool.Data.SHClassRecord> classList = SHSchool.Data.SHClass.SelectByIDs(classIDList);
            //班級ID對照清單
            Dictionary<string,SHSchool.Data.SHClassRecord> classDic = new Dictionary<string,Data.SHClassRecord>();
            foreach (SHSchool.Data.SHClassRecord classRecord in classList)
            {
                if (!classDic.ContainsKey(classRecord.ID))
                {
                    classDic.Add(classRecord.ID, classRecord);
                }
            }

            List<StudentSortObj_SHSchoolData> list = new List<StudentSortObj_SHSchoolData>();
            foreach (SHSchool.Data.SHStudentRecord student in StudentList)
            {
                if (!string.IsNullOrEmpty(student.RefClassID))
                {
                    StudentSortObj_SHSchoolData obj = new StudentSortObj_SHSchoolData(classDic[student.RefClassID],student);
                    list.Add(obj);
                }
                else
                {
                    StudentSortObj_SHSchoolData obj = new StudentSortObj_SHSchoolData(null, student);
                    list.Add(obj);
                }
            }
            list.Sort(SortSHSchoolData_SHStudentRecord);

            return list.Select(x => x._StudentRecord).ToList();

        }

        static private int SortSHSchoolData_SHStudentRecord(StudentSortObj_SHSchoolData obj1, StudentSortObj_SHSchoolData obj2)
        {
            return obj1._SortString.CompareTo(obj2._SortString);
        }
        
        #endregion

        #region K12ClassRecord

        static public List<K12.Data.ClassRecord> K12Data_ClassRecord(List<K12.Data.ClassRecord> ClassList)
        {
            ClassList.Sort(SortK12Data_ClassRecord);
            return ClassList;
        }

        static private int SortK12Data_ClassRecord(K12.Data.ClassRecord class1, K12.Data.ClassRecord class2)
        {
            string ClassYear1 = class1.GradeYear.HasValue ? class1.GradeYear.Value.ToString().PadLeft(10, '0') : string.Empty.PadLeft(10, '9');
            string ClassYear2 = class2.GradeYear.HasValue ? class2.GradeYear.Value.ToString().PadLeft(10, '0') : string.Empty.PadLeft(10, '9');

            string DisplayOrder1 = "";
            if (string.IsNullOrEmpty(class1.DisplayOrder))
            {
                DisplayOrder1 = class1.DisplayOrder.PadLeft(10, '9');
            }
            else
            {
                DisplayOrder1 = class1.DisplayOrder.PadLeft(10, '0');
            }
            string DisplayOrder2 = "";
            if (string.IsNullOrEmpty(class2.DisplayOrder))
            {
                DisplayOrder2 = class2.DisplayOrder.PadLeft(10, '9');
            }
            else
            {
                DisplayOrder2 = class2.DisplayOrder.PadLeft(10, '0');
            }

            string ClassName1 = class1.Name.PadLeft(10, '0');
            string ClassName2 = class2.Name.PadLeft(10, '0');

            string Compareto1 = ClassYear1 + DisplayOrder1 + ClassName1;
            string Compareto2 = ClassYear2 + DisplayOrder2 + ClassName2;

            return Compareto1.CompareTo(Compareto2);
        }

        static public List<K12.Data.StudentRecord> K12Data_StudentRecord(List<K12.Data.StudentRecord> StudentList)
        {
            //整理出學生&班級資料清單
            List<string> classIDList = new List<string>();
            foreach (K12.Data.StudentRecord student in StudentList)
            {
                if (!string.IsNullOrEmpty(student.RefClassID))
                {
                    if (!classIDList.Contains(student.RefClassID))
                    {
                        classIDList.Add(student.RefClassID);
                    }
                }
            }
            //一次取得班級清單
            List<K12.Data.ClassRecord> classList = K12.Data.Class.SelectByIDs(classIDList);
            //班級ID對照清單
            Dictionary<string, K12.Data.ClassRecord> classDic = new Dictionary<string, K12.Data.ClassRecord>();
            foreach (K12.Data.ClassRecord classRecord in classList)
            {
                if (!classDic.ContainsKey(classRecord.ID))
                {
                    classDic.Add(classRecord.ID, classRecord);
                }
            }

            List<StudentSortObj_K12Data> list = new List<StudentSortObj_K12Data>();
            foreach (K12.Data.StudentRecord student in StudentList)
            {
                if (!string.IsNullOrEmpty(student.RefClassID))
                {
                    StudentSortObj_K12Data obj = new StudentSortObj_K12Data(classDic[student.RefClassID], student);
                    list.Add(obj);
                }
                else
                {
                    StudentSortObj_K12Data obj = new StudentSortObj_K12Data(null, student);
                    list.Add(obj);
                }
            }
            list.Sort(SortK12Data_StudentRecord);

            return list.Select(x => x._StudentRecord).ToList();

        }

        static private int SortK12Data_StudentRecord(StudentSortObj_K12Data obj1, StudentSortObj_K12Data obj2)
        {
            return obj1._SortString.CompareTo(obj2._SortString);
        }
        #endregion
    }
}
