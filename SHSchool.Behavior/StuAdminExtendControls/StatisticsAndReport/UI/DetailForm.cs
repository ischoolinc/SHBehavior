using System;
using System.Collections.Generic;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using FISCA.Presentation.Controls;

namespace SHSchool.Behavior.StuAdminExtendControls
{
    internal partial class DetailForm : BaseForm
    {
        private CellGradeYear _parent;

        private StudentCollection _students;
        public StudentCollection Students
        {
            get { return _students; }
            set { _students = value; }
        }

        public DetailForm(CellGradeYear parent)
        {
            InitializeComponent();
            _parent = parent;
            InitialMenuBarItem();
        }

        private void DetailForm_Load(object sender, EventArgs e)
        {
            FillData();
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

            if (dataGridViewX1.SelectedRows.Count > 0)
            {
                List<Student> list = new List<Student>();
                foreach (DataGridViewRow row in dataGridViewX1.SelectedRows)
                    list.Add(row.Tag as Student);

                _parent.SetCategory(list, (sender as ButtonItem).Tag.ToString());
                RefreshStudents();
            }
        }

        private void FillData()
        {
            dataGridViewX1.SuspendLayout();
            dataGridViewX1.Rows.Clear();
            foreach (Student stu in _students.Values)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1, stu.Department, stu.GradeYear, stu.ClassName, stu.StudentNumber, stu.Name, stu.Gender, stu.Location);
                row.Tag = stu;
                dataGridViewX1.Rows.Add(row);
            }
            dataGridViewX1.ResumeLayout();
        }

        public void RefreshStudents()
        {
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                row.Cells[colCate.Name].Value = (row.Tag as Student).Location;
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}