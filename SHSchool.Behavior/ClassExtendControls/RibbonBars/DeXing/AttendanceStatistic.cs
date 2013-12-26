using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using FISCA.DSAUtil;
using DevComponents.DotNetBar;
using System.Xml;
using Aspose.Cells;
using System.IO;
using System.Diagnostics;
using SmartSchool.Common;
using Framework.Feature;
using K12.Data;
using SHSchool.Behavior.StudentExtendControls;

namespace SHSchool.Behavior.ClassExtendControls
{
    public partial class AttendanceStatistic : UserControl, IDeXingExport
    {
        private string[] _classidList;
        public AttendanceStatistic(string[] classidList)
        {
            InitializeComponent();
            _classidList = classidList;
        }

        #region IDeXingExport 成員

        public void LoadData()
        {
            DSResponse dsrsp = Config.GetAbsenceList();
            DSXmlHelper helper = dsrsp.GetContent();
            foreach (XmlElement e in helper.GetElements("Absence"))
            {
                string name = e.GetAttribute("Name");
                int rowIndex = dataGridView.Rows.Add();
                DataGridViewRow row = dataGridView.Rows[rowIndex];
                row.Cells[1].Value = name;
            }
        }

        public void Export()
        {
            dataGridView.EndEdit();
            if (!IsValid())
            {
                MsgBox.Show("輸入資料有誤，請修正後再進行匯出！", "驗證錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement(".", "SchoolYear", School.DefaultSchoolYear);
            helper.AddElement(".", "Semester", School.DefaultSemester);

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (!IsCheckedRow(row)) continue;
                XmlElement e = helper.AddElement("Absence");
                e.SetAttribute("Name", row.Cells[1].Value.ToString());
                e.SetAttribute("PeriodCount", row.Cells[2].Value.ToString());
            }

            foreach (string id in _classidList)
            {
                helper.AddElement(".", "ClassID", id);
            }
            DSResponse dsrsp = QueryAttendance.GetAttendanceStatistic(new DSRequest(helper));
            if (!dsrsp.HasContent)
            {
                MsgBox.Show("取得回覆資料失敗:" + dsrsp.GetFault().Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DSXmlHelper rsp = dsrsp.GetContent();
            Workbook book = new Workbook();
            string A1Name = string.Empty;
            book.Worksheets.Clear();
            foreach (XmlElement e in rsp.GetElements("Absence"))
            {

                int sheetIndex = book.Worksheets.Add();
                Worksheet sheet = book.Worksheets[sheetIndex];
                sheet.Name = e.GetAttribute("Type");

                string schoolName = School.ChineseName;
                Cell A1 = sheet.Cells["A1"];
                A1.Style.Borders.SetColor(Color.Black);
                A1Name = "扣分累計學生清單(" + e.GetAttribute("Type") + ")";
                A1.PutValue(A1Name);
                A1.Style.HorizontalAlignment = TextAlignmentType.Center;
                sheet.Cells.Merge(0, 0, 1, 5);

                FormatCell(sheet.Cells["A2"], "班級");
                FormatCell(sheet.Cells["B2"], "座號");
                FormatCell(sheet.Cells["C2"], "姓名");
                FormatCell(sheet.Cells["D2"], "學號");
                FormatCell(sheet.Cells["E2"], "累積節次");
                //FormatCell(sheet.Cells["F2"], "累積扣分");

                int index = 3;
                foreach (XmlElement s in e.SelectNodes("Student"))
                {
                    FormatCell(sheet.Cells["A" + index], s.GetAttribute("ClassName"));
                    FormatCell(sheet.Cells["B" + index], s.GetAttribute("SeatNo"));
                    FormatCell(sheet.Cells["C" + index], s.GetAttribute("Name"));
                    FormatCell(sheet.Cells["D" + index], s.GetAttribute("StudentNumber"));
                    FormatCell(sheet.Cells["E" + index], s.GetAttribute("PeriodCount"));
                    //FormatCell(sheet.Cells["F" + index], s.GetAttribute("Subtract"));
                    index++;
                }
            }

            string path = Path.Combine(Application.StartupPath, "Reports");

            //如果目錄不存在則建立。
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            path = Path.Combine(path, "缺曠累計名單" + ".xls");
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
            try
            {
                book.Save(path);
            }
            catch (IOException)
            {
                try
                {
                    FileInfo file = new FileInfo(path);
                    string nameTempalte = file.FullName.Replace(file.Extension, "") + "{0}.xls";
                    int count = 1;
                    string fileName = string.Format(nameTempalte, count);
                    while (File.Exists(fileName))
                        fileName = string.Format(nameTempalte, count++);

                    book.Save(fileName);
                    path = fileName;
                }
                catch (Exception ex)
                {
                    MsgBox.Show("檔案儲存失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show("檔案儲存失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Process.Start(path);
            }
            catch (Exception ex)
            {
                MsgBox.Show("檔案開啟失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        public Control MainControl
        {
            get { return this.groupPanel1; }
        }

        private string ConvertToValidName(string A1Name)
        {
            char[] invalids = Path.GetInvalidFileNameChars();

            string result = A1Name;
            foreach (char each in invalids)
                result = result.Replace(each, '_');

            return result;
        }

        private void FormatCell(Cell cell, string value)
        {
            cell.PutValue(value);
            cell.Style.Borders.SetStyle(CellBorderType.Hair);
            cell.Style.Borders.SetColor(Color.Black);
            cell.Style.Borders.DiagonalStyle = CellBorderType.None;
            cell.Style.HorizontalAlignment = TextAlignmentType.Center;
        }

        private bool IsCheckedRow(DataGridViewRow row)
        {
            if (row.Cells[0].Value == null)
                return false;
            string value = row.Cells[0].Value.ToString();
            bool check = false;
            if (!bool.TryParse(value, out check))
                return false;
            return check;
        }
        #endregion

        //private void dataGridView_RowValidated(object sender, DataGridViewCellEventArgs e)
        //{
            
        //}

        private bool IsValid()
        {
            errorProvider1.Clear();
            bool valid = true;
            int count = 0;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (IsCheckedRow(row))
                {
                    if (row.Cells[2].ErrorText == string.Empty)
                        count++;
                    else
                        valid = false;
                }
            }
            if (count == 0)
            {
                errorProvider1.SetError(dataGridView, "至少必須選擇一個缺曠類別");                
                return false;
            }
            return valid;
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView.Rows[e.RowIndex];
            //DataGridViewCheckBoxCell cell = row.Cells[0] as DataGridViewCheckBoxCell;
            //if (cell.Value == null)
            //    return;

            DataGridViewCell c = row.Cells[2];
            c.ErrorText = string.Empty;
            string value = c.Value == null ? "0" : c.Value.ToString();
            decimal subtract = 0;
            if (!decimal.TryParse(value, out subtract))
            {
                row.Cells[2].ErrorText = "必須為數字";
                return;
            }
            c.Value = subtract;
            row.Cells[2].ErrorText = string.Empty;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            StatisticWeightConfig config = new StatisticWeightConfig();
            config.ShowDialog();
        }
    }
}
