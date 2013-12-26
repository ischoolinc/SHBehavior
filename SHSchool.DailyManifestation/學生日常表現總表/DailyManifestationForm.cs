using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;

namespace SHSchool.DailyManifestation
{
    public partial class DailyManifestationForm : BaseForm
    {

        //string _StudentID { get; set; }

        //BackgroundWorker BGW = new BackgroundWorker();

        public DailyManifestationForm(string StudentID)
        {
            InitializeComponent();
            //_StudentID = StudentID;
            //BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            //BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            //BGW.ProgressChanged += new ProgressChangedEventHandler(BGW_ProgressChanged);
        }

        //void BGW_DoWork(object sender, DoWorkEventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

        //void BGW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        //{
        //    FISCA.Presentation.MotherForm.SetStatusBarMessage("", e.ProgressPercentage);
        //}

        //void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        //{
        //    //列印完成
        //}

        //private void LockForm
        //{
        //    set
        //    {

        //    }
        //}

        private void DailyManifestationForm_Load(object sender, EventArgs e)
        {
            //本報表為學生每學期之資料總表



            ////獎

            //List<MeritRecord> MeritList = Merit.SelectByStudentIDs(new List<string> { _StudentID });

            ////懲
            //List<DemeritRecord> DemeritList = Demerit.SelectByStudentIDs(new List<string> { _StudentID });

            ////缺曠
            //List<AttendanceRecord> AttendanceList = Attendance.SelectByStudentIDs(new List<string> { _StudentID });

            ////日常生活表現
            //List<MoralScoreRecord> MoralScoreList = MoralScore.SelectByStudentIDs(new List<string> { _StudentID });

            ////異動記錄
            //List<UpdateRecordRecord> UpdateRecordList = UpdateRecord.SelectByStudentIDs(new List<string> { _StudentID });









        }
    }
}
