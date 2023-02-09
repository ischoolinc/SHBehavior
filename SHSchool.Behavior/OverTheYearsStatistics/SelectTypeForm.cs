using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using FISCA.DSAUtil;
using SmartSchool.Common;
using K12.Data.Configuration;

namespace SHSchool.Behavior
{
    public partial class SelectTypeForm : BaseForm
    {
        private string _preferenceElementName;
        private string _configDefName = "節次分類";

        private BackgroundWorker _BGWAbsenceAndPeriodList;

        private List<string> typeList = new List<string>();
        private List<string> absenceList = new List<string>();
       
        public SelectTypeForm(string name)
        {
            InitializeComponent();

            _preferenceElementName = name;

            _BGWAbsenceAndPeriodList = new BackgroundWorker();
            _BGWAbsenceAndPeriodList.DoWork += new DoWorkEventHandler(_BGWAbsenceAndPeriodList_DoWork);
            _BGWAbsenceAndPeriodList.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWAbsenceAndPeriodList_RunWorkerCompleted);
            _BGWAbsenceAndPeriodList.RunWorkerAsync();
        }

        void _BGWAbsenceAndPeriodList_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;

            System.Windows.Forms.DataGridViewTextBoxColumn colName = new DataGridViewTextBoxColumn();
            colName.HeaderText = "節次分類";
            colName.MinimumWidth = 70;
            colName.Name = "colName";
            colName.ReadOnly = true;
            colName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            colName.Width = 70;
            this.dataGridViewX1.Columns.Add(colName);

            foreach (string absence in absenceList)
            {
                System.Windows.Forms.DataGridViewCheckBoxColumn newCol = new DataGridViewCheckBoxColumn();
                newCol.HeaderText = absence;
                newCol.Width = 55;
                newCol.ReadOnly = false;
                newCol.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
                newCol.Tag = absence;
                newCol.ValueType = typeof(bool);
                this.dataGridViewX1.Columns.Add(newCol);
            }

            foreach (string type in typeList)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1, type);
                row.Tag = type;
                dataGridViewX1.Rows.Add(row);
            }


            #region 讀取列印設定 Preference

            K12.Data.Configuration.ConfigData Defconfig = K12.Data.School.Configuration[_preferenceElementName];

            XmlElement config;

            if (Defconfig.GetXml(_configDefName, null) == null)
            {
                config = new XmlDocument().CreateElement(_configDefName);
                Defconfig[_configDefName] = config.OuterXml;
                Defconfig.Save();
            }
            else
            {
                #region 已有設定檔則將設定檔內容填回畫面上

                config = Defconfig.GetXml(_configDefName, null);

                foreach (XmlElement type in config.SelectNodes("Type"))
                {
                    string typeName = type.GetAttribute("Text");
                    foreach (DataGridViewRow row in dataGridViewX1.Rows)
                    {
                        if (typeName == ("" + row.Tag))
                        {
                            foreach (XmlElement absence in type.SelectNodes("Absence"))
                            {
                                string absenceName = absence.GetAttribute("Text");
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    if (cell.OwningColumn is DataGridViewCheckBoxColumn && ("" + cell.OwningColumn.Tag) == absenceName)
                                    {
                                        cell.Value = true;
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
                #endregion
            }

            #endregion
        }

        void _BGWAbsenceAndPeriodList_DoWork(object sender, DoWorkEventArgs e)
        {
            typeList = GetConfig.GetPeriodList();

            absenceList = GetConfig.GetAbsenceList();

        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (!CheckColumnNumber())
                return;

            #region 更新列印設定 Preference

            K12.Data.Configuration.ConfigData Defconfig = K12.Data.School.Configuration[_preferenceElementName];

            XmlElement config;

            if (Defconfig.GetXml(_configDefName, null) == null)
            {
                config = new XmlDocument().CreateElement(_configDefName);
            }
            else
            {
                config = Defconfig.GetXml(_configDefName, null);
            }

            config.RemoveAll();

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                bool needToAppend = false;
                XmlElement type = config.OwnerDocument.CreateElement("Type");
                type.SetAttribute("Text", "" + row.Tag);
                foreach (DataGridViewCell cell in row.Cells)
                {
                    XmlElement absence = config.OwnerDocument.CreateElement("Absence");
                    absence.SetAttribute("Text", "" + cell.OwningColumn.Tag);
                    if (cell.Value is bool && ((bool)cell.Value))
                    {
                        needToAppend = true;
                        type.AppendChild(absence);
                    }
                }
                if (needToAppend)
                    config.AppendChild(type);
            }

            //foreach (TreeNode typeNode in treeView1.Nodes)
            //{
            //    XmlElement type = config.OwnerDocument.CreateElement("Type");
            //    type.SetAttribute("Text", typeNode.Text);
            //    type.SetAttribute("Checked", typeNode.Checked.ToString());

            //    foreach (TreeNode absenceNode in typeNode.Nodes)
            //    {
            //        if (absenceNode.Checked == true)
            //        {
            //            XmlElement absence = config.OwnerDocument.CreateElement("Absence");
            //            absence.SetAttribute("Text", absenceNode.Text);
            //            type.AppendChild(absence);
            //        }
            //    }
            //    config.AppendChild(type);
            //}

            Defconfig[_configDefName] = config.OuterXml;
            Defconfig.Save();

            #endregion

            this.DialogResult = DialogResult.OK;
        }

        internal bool CheckColumnNumber()
        {
            int limit = 253;
            int columnNumber = 0;
            int block = 9;

            //foreach (TreeNode type in treeView1.Nodes)
            //{
            //    foreach (TreeNode var in type.Nodes)
            //    {
            //        if (var.Checked == true)
            //            columnNumber++;
            //    }
            //}
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value is bool && ((bool)cell.Value))
                        columnNumber++;
                }
            }

            if (columnNumber * block > limit)
            {
                MsgBox.Show("您所選擇的假別超出 Excel 的最大欄位，請減少部分假別");
                return false;
            }
            else
                return true;
        }

        private void dataGridViewX1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ColumnIndex == 0)
                        continue;

                    cell.Value = checkBoxX1.Checked;
                }
            }
        }
    }
}