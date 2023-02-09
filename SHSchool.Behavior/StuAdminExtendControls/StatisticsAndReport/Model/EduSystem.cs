using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    internal class EduSystem
    {
        private Workbook _workbook;
        private Worksheet _worksheet;

        private Dictionary<string, int> _row_index_table;
        private Dictionary<string, int> _col_index_table;

        private Dictionary<string, int> RowTable
        {
            get { return _row_index_table; }
        }
        private Dictionary<string, int> ColTable
        {
            get { return _col_index_table; }
        }

        private string _school_year;
        public string SchoolYear
        {
            get { return _school_year; }
            set { _school_year = value; }
        }

        private string _semester;
        public string Semester
        {
            get { return _semester; }
            set { _semester = value; }
        }

        private Dictionary<string, RowDiscipline> _rows;

        private string _name;
        public string Name
        {
            get { return _name; }
        }

        public Worksheet Worksheet
        {
            get { return _worksheet; }
        }

        public EduSystem(Workbook workbook, string name)
        {
            _workbook = workbook;
            _name = name;
            _worksheet = _workbook.Worksheets[_workbook.Worksheets.Add()];
            _rows = new Dictionary<string, RowDiscipline>();

            _row_index_table = new Dictionary<string, int>();
            _col_index_table = new Dictionary<string, int>();

            LoadTemplate();
            CreateRowDiscipline();
        }

        private void LoadTemplate()
        {
            Workbook template = new Workbook();
            template.Open(new MemoryStream(Properties.Resources.獎懲人數統計報表));

            switch (_name)
            {
                case "普通科":
                case "綜合高中":
                    _worksheet.Copy(template.Worksheets[0]);
                    BuildIndexTable(3);
                    break;
                case "職業科":
                    _worksheet.Copy(template.Worksheets[1]);
                    BuildIndexTable(3);
                    break;
                default:
                    throw new Exception("沒有這個東西");
            }

            _worksheet.Name = _name;
        }

        private void BuildIndexTable(int year)
        {
            int row_start = 7;
            int row_end = 21;
            for (int i = row_start; i <= row_end; i++)
            {
                string str = _worksheet.Cells[i, 0].StringValue;
                if (!string.IsNullOrEmpty(str))
                {
                    str = str.Replace(" ", "");
                    str = str.Replace("(一)", "");
                    str = str.Replace("(二)", "");
                    str = str.Replace("(三)", "");
                    _row_index_table.Add(str, i);
                }
            }

            int col_start = 5;
            for (int i = 2; i <= _worksheet.Cells.MaxColumn; i++)
            {
                string cate = "";
                if (!string.IsNullOrEmpty(_worksheet.Cells[col_start, i].StringValue))
                    cate = _worksheet.Cells[col_start, i].StringValue;
                else if (!string.IsNullOrEmpty(_worksheet.Cells[col_start, i - 1].StringValue))
                    cate = _worksheet.Cells[col_start, i - 1].StringValue;
                else if (!string.IsNullOrEmpty(_worksheet.Cells[col_start, i - 2].StringValue))
                    cate = _worksheet.Cells[col_start, i - 2].StringValue;

                string subcate = _worksheet.Cells[col_start + 1, i].StringValue;

                if (!string.IsNullOrEmpty(cate) && !string.IsNullOrEmpty(subcate))
                {
                    _col_index_table.Add(cate.Replace(" ", "") + "/" + subcate.Replace(" ", ""), i);
                }
            }
        }

        private void CreateRowDiscipline()
        {
            _rows.Add("獎勵", new RowDiscipline("獎勵"));
            _rows.Add("嘉獎", new RowDiscipline("嘉獎"));
            _rows.Add("小功", new RowDiscipline("小功"));
            _rows.Add("大功", new RowDiscipline("大功"));
            _rows.Add("懲罰", new RowDiscipline("懲罰"));
            _rows.Add("警告", new RowDiscipline("警告"));
            _rows.Add("小過", new RowDiscipline("小過"));
            _rows.Add("大過", new RowDiscipline("大過"));
            _rows.Add("留校察看", new RowDiscipline("留校察看"));
        }

        public void AddRecord(Student student)
        {
            int award = student.AwardA + student.AwardB + student.AwardC;
            int fault = student.FaultA + student.FaultB + student.FaultC;

            if (award > 0)
            {
                if (student.AwardC > 0)
                    _rows["嘉獎"].AddRecord(student, student.AwardC);
                if (student.AwardB > 0)
                    _rows["小功"].AddRecord(student, student.AwardB);
                if (student.AwardA > 0)
                    _rows["大功"].AddRecord(student, student.AwardA);

                _rows["獎勵"].AddRecord(student, 0);
            }

            if (fault > 0)
            {
                if (student.FaultC > 0)
                    _rows["警告"].AddRecord(student, student.FaultC);
                if (student.FaultB > 0)
                    _rows["小過"].AddRecord(student, student.FaultB);
                if (student.FaultA > 0)
                    _rows["大過"].AddRecord(student, student.FaultA);

                _rows["懲罰"].AddRecord(student, 0);
            }

            if (student.UltimateAdmonition)
                _rows["留校察看"].AddRecord(student, 1);
        }

        public void FillData()
        {
            foreach (string kind in _rows.Keys)
            {
                foreach (string key in _rows[kind].Records.Keys)
                {
                    if (!RowTable.ContainsKey(kind))
                        continue;
                    if (!ColTable.ContainsKey(key))
                        continue;

                    int rowIndex = RowTable[kind];
                    int colIndex = ColTable[key];

                    if (kind == "獎勵" || kind == "懲罰")
                    {
                        if (_rows[kind].Records[key].StudentCount > 0)
                            _worksheet.Cells[rowIndex, colIndex].PutValue(_rows[kind].Records[key].StudentCount);
                    }
                    else
                    {
                        if (_rows[kind].Records[key].StudentCount > 0)
                            _worksheet.Cells[rowIndex, colIndex].PutValue(_rows[kind].Records[key].StudentCount);
                        if (_rows[kind].Records[key].Times > 0)
                            _worksheet.Cells[rowIndex + 1, colIndex].PutValue(_rows[kind].Records[key].Times);
                    }
                }
            }

            //填入學校代碼
            _worksheet.Cells[4, 0].PutValue(SmartSchool.Customization.Data.SystemInformation.SchoolCode);

            //填入報表標題
            _worksheet.Cells[2, 0].PutValue("高中高職學校學生獎懲人數—" + _name);

            //填入學年度學期
            _worksheet.Cells[3, 2].PutValue("中華民國         " + _school_year + " 學年  第 " + _semester + " 學期");
        }
    }
}
