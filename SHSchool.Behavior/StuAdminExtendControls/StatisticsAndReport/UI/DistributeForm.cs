using System;
using System.Collections.Generic;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using FISCA.Presentation.Controls;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    internal partial class DistributeForm : BaseForm
    {
        private StudentCollection _students;
        public StudentCollection Students
        {
            get
            {
                if (_students != null)
                    return _students;
                return new StudentCollection();
            }
        }

        Dictionary<string, RowDept> _dept_list;

        public DistributeForm(StudentCollection students, string school_year, string semester)
        {
            InitializeComponent();

            _students = students;
            this.Text += "  " + school_year + " 學年度  ";
            this.Text += "第 " + semester + " 學期";
            InitialMenuBarItem();
        }

        private void InitialMenuBarItem()
        {
            List<string> cates = new List<string>(new string[] {
                "普通科/一年級",
                "普通科/二年級",
                "普通科/三年級",
                "普通科/延修生",
                "職業科/一年級",
                "職業科/二年級",
                "職業科/三年級",
                "職業科/四年級",
                "職業科/延修生",
                "綜合高中/一年級",
                "綜合高中/二年級",
                "綜合高中/三年級",
                "綜合高中/延修生",
            });

            Dictionary<string, BaseItem> buttons = new Dictionary<string, BaseItem>();

            foreach (string each in cates)
            {
                string cate = each.Split('/')[0];
                string year = each.Split('/')[1];

                if (!buttons.ContainsKey(cate))
                {
                    ButtonItem item = new ButtonItem();
                    item.Text = "統計為「" + cate + "」";
                    item.Tag = cate;
                    item.Click += new EventHandler(Item_Click);

                    itemContainer1.SubItems.Add(item);

                    ItemContainer newContainer = new ItemContainer();
                    newContainer.LayoutOrientation = DevComponents.DotNetBar.eOrientation.Vertical;
                    item.SubItems.Add(newContainer);
                    buttons.Add(cate, newContainer);
                }

                ButtonItem subitem = new ButtonItem();
                subitem.Text = year;
                subitem.Tag = each;
                subitem.Click += new EventHandler(Item_Click);
                buttons[cate].SubItems.Add(subitem);
            }

            ButtonItem clear_item = new ButtonItem();
            clear_item.Text = "清除設定";
            clear_item.Tag = "";
            clear_item.Click += new EventHandler(Item_Click);
            itemContainer1.SubItems.Add(clear_item);
        }

        void Item_Click(object sender, EventArgs e)
        {
            if ((sender as ButtonItem).SubItems.Count > 0)
            {
                buttonItem1.ClosePopup();
            }

            if (dataGridViewX1.SelectedCells.Count > 0)
            {
                foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                {
                    if (cell is CellGradeYear && cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        CellGradeYear cellGradeYear = cell as CellGradeYear;
                        cellGradeYear.SetCategory(new List<Student>(cellGradeYear.Students.Values), (sender as ButtonItem).Tag.ToString());
                    }
                }
                dataGridViewX1.Refresh();
            }
        }

        private void dataGridViewX1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.RowIndex < 0)
                return;

            DataGridViewCell cell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            if (cell is CellGradeYear && cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
            {
                CellGradeYear cellGradeYear = cell as CellGradeYear;
                cellGradeYear.Detail.ShowDialog();
                dataGridViewX1.Refresh();
            }
        }

        private void DistributeForm_Load(object sender, EventArgs e)
        {
            _dept_list = new Dictionary<string, RowDept>();

            foreach (Student stu in _students.Values)
            {
                string dept = stu.Department;
                if (string.IsNullOrEmpty(dept))
                    dept = "沒有科別";

                if (!_dept_list.ContainsKey(dept))
                    _dept_list.Add(dept, new RowDept(dataGridViewX1, dept));

                _dept_list[dept].AddStudent(stu);
            }

            List<string> sorted = new List<string>();
            sorted.AddRange(_dept_list.Keys);
            sorted.Sort();
            if (sorted.Contains("沒有科別"))
            {
                sorted.Remove("沒有科別");
                sorted.Insert(sorted.Count, "沒有科別");
            }

            dataGridViewX1.SuspendLayout();
            foreach (string key in sorted)
            {
                dataGridViewX1.Rows.Add(_dept_list[key]);
            }
            dataGridViewX1.ResumeLayout();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (MessageBoxEx.Show("您確定要離開？", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                this.Close();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            //學生有設定分類才進行動作
            foreach (Student stu in _students.Values)
            {
                if (!string.IsNullOrEmpty(stu.Location))
                {
                    this.DialogResult = DialogResult.OK;
                    break;
                }
            }
        }
    }
}