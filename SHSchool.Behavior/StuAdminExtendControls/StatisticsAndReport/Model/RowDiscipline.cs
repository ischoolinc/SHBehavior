using System.Collections.Generic;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    /// <summary>
    /// 報表上每一個獎懲類別
    /// </summary>
    internal class RowDiscipline
    {
        private Dictionary<string, Record> _records;
        public Dictionary<string, Record> Records
        {
            get { return _records; }
        }

        private string _name;
        public string Name
        {
            get { return _name; }
        }

        public RowDiscipline(string name)
        {
            _name = name;
            _records = new Dictionary<string, Record>();
            _records.Add("總計/男", new Record());
            _records.Add("總計/女", new Record());
            _records.Add("總計/計", new Record());
        }

        public void AddRecord(Student student, int times)
        {
            string gg = student.Location.Split('/')[1] + "/" + student.Gender;

            if (!_records.ContainsKey(gg))
                _records.Add(gg, new Record());
            _records[gg].StudentCount++;
            _records[gg].Times += times;

            if (student.Gender == "男")
            {
                _records["總計/男"].StudentCount++;
                _records["總計/男"].Times += times;
            }
            else if (student.Gender == "女")
            {
                _records["總計/女"].StudentCount++;
                _records["總計/女"].Times += times;
            }

            _records["總計/計"].StudentCount++;
            _records["總計/計"].Times += times;
        }
    }

    /// <summary>
    /// 報表上每一筆小紀錄，包含人數及次數
    /// </summary>
    internal class Record
    {
        public Record()
        {
        }

        private int _student_count = 0;
        public int StudentCount
        {
            get { return _student_count; }
            set { _student_count = value; }
        }

        private int _times = 0;
        public int Times
        {
            get { return _times; }
            set { _times = value; }
        }

        //public Record AddRecord(Record rec)
        //{
        //    _student_count += rec.StudentCount;
        //    _times += rec.Times;

        //    return this;
        //}
    }
}
