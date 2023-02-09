using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Aspose.Cells;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    internal class ReportGenerator
    {
        private StudentCollection _students;
        private string _school_year;
        private string _semester;

        private Workbook _workbook;
        private Dictionary<string, EduSystem> _systems;

        public ReportGenerator(StudentCollection students, string school_year, string semester)
        {
            _students = students;
            _school_year = school_year;
            _semester = semester;

            _workbook = new Workbook();
            _workbook.Open(new MemoryStream(Properties.Resources.獎懲人數統計報表));
            _workbook.Worksheets.Clear();
            _systems = new Dictionary<string, EduSystem>();

            foreach (string cate in new string[] { "普通科", "職業科", "綜合高中" })
            {
                _systems.Add(cate, new EduSystem(_workbook, cate));
            }

            ProcessDiscipline();

            foreach (EduSystem sys in _systems.Values)
            {
                sys.SchoolYear = _school_year;
                sys.Semester = _semester;
                sys.FillData();
            }
        }

        private void ProcessDiscipline()
        {
            foreach (Student stu in _students.Values)
            {
                if (string.IsNullOrEmpty(stu.Location))
                    continue;

                string cate = stu.Location.Split('/')[0];
                _systems[cate].AddRecord(stu);
            }
        }

        public string Generate()
        {
            _workbook.Worksheets.ActiveSheetIndex = 0;
            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, "獎懲人數統計表.xlsx");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                _workbook.Save(path);
            }
            catch
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = "獎懲人數統計表.xlsx";
                sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _workbook.Save(sd.FileName);
                        path = sd.FileName;
                    }
                    catch
                    {
                        MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return "";
                    }
                }
            }

            return path;
        }
    }
}
