using System;
using System.Collections.Generic;
using System.Text;
using FISCA.DSAUtil;
using FISCA.Authentication;

namespace SHSchool.Behavior.StudentExtendControls
{
    public class EditAttendance2
    {
        [AutoRetryOnWebException()]
        public static void Delete(DSRequest dSRequest)
        {
            DSAServices.CallService("SmartSchool.Student.Attendance.Delete", dSRequest);
        }

        [AutoRetryOnWebException()]
        public static void Update(DSRequest dSRequest)
        {
            DSAServices.CallService("SmartSchool.Student.Attendance.Update", dSRequest);
        }

        public static void Insert(DSRequest dSRequest)
        {
            DSAServices.CallService("SmartSchool.Student.Attendance.Insert", dSRequest);
        }
    }
}
