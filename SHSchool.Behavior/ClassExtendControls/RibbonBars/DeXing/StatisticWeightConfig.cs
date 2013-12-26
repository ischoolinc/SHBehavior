using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SmartSchool.Common;
using FISCA.DSAUtil;
using System.Xml;
using FISCA.Presentation.Controls;
using Framework.Feature;
using SHSchool.Behavior.Feature;

namespace SHSchool.Behavior.ClassExtendControls
{
    public partial class StatisticWeightConfig : FISCA.Presentation.Controls.BaseForm
    {
        public StatisticWeightConfig()
        {
            InitializeComponent();

            DSResponse dsrsp = Config.GetPeriodList();
            DSXmlHelper helper = dsrsp.GetContent();
            List<PeriodInfo> collection = new List<PeriodInfo>();
            foreach (XmlElement element in helper.GetElements("Period"))
            {
                PeriodInfo info = new PeriodInfo(element);
                collection.Add(info);
            }
            collection.Sort(SortByOrder);
            foreach (PeriodInfo info in collection)
            {
                int index = dataGridView.Rows.Add();
                DataGridViewRow row = dataGridView.Rows[index];
                row.Cells[colPeriod.Index].Value = info.Name;
                row.Cells[colType.Index].Value = info.Type;
                row.Cells[colOrder.Index].Value = info.Sort;
                row.Cells[colWeight.Index].Value = info.Aggregated;

                ValidateRow(row);
            }
        }

        private static int SortByOrder(PeriodInfo info1, PeriodInfo info2)
        {
            return info1.Sort.CompareTo(info2.Sort);
        }

        private bool ValidateRow(DataGridViewRow row)
        {
            bool pass = true;
            if (row.IsNewRow)
                return true;
            //不得空白
            #region 不得空白
            foreach (DataGridViewCell cell in row.Cells)
            {
                if ("" + cell.Value == "")
                {
                    cell.ErrorText = "不得空白";
                    pass &= false;
                    dataGridView.UpdateCellErrorText(cell.ColumnIndex, row.Index);
                }
                else if (cell.ErrorText == "不得空白")
                {
                    cell.ErrorText = "";
                    dataGridView.UpdateCellErrorText(cell.ColumnIndex, row.Index);
                }
            }
            #endregion
            //不得重複
            #region 不得重複
            foreach (DataGridViewRow r in dataGridView.Rows)
            {
                if (r != row)
                {
                    foreach (int index in new int[] { colPeriod.Index })
                    {
                        if ("" + r.Cells[index].Value == "" + row.Cells[index].Value)
                        {
                            row.Cells[index].ErrorText = "不得重複";
                            dataGridView.UpdateCellErrorText(index, row.Index);
                            pass &= false;
                        }
                        else if (row.Cells[index].ErrorText == "不得重複")
                        {
                            row.Cells[index].ErrorText = "";
                            dataGridView.UpdateCellErrorText(index, row.Index);
                        }
                    }
                }
            }
            #endregion
            
            //節次對照必須為數值
            #region 節次對照必須為數值
            decimal dec;
            if (!decimal.TryParse("" + row.Cells[colWeight.Index].Value, out dec) || dec < 0)
            {
                row.Cells[colWeight.Index].ErrorText = "必須輸入０或正數";
                dataGridView.UpdateCellErrorText(colWeight.Index, row.Index);
                pass &= false;
            }
            else if (row.Cells[colWeight.Index].ErrorText == "必須輸入０或正數")
            {
                row.Cells[colWeight.Index].ErrorText = "";
                dataGridView.UpdateCellErrorText(colWeight.Index, row.Index);
            }
            #endregion
            return pass;
        }

        private void dataGridView_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView.SelectedCells.Count == 1)
                dataGridView.BeginEdit(true);
        }

        private void dataGridView_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView.EndEdit();
        }

        private void dataGridView_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
            ValidateRow(dataGridView.Rows[e.RowIndex]);
        }

        private void dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.ColumnIndex == colWeight.Index)
            //{
            //    foreach (DataGridViewRow row in dataGridView.Rows)
            //    {
            //        if (row.Index == e.RowIndex || row.IsNewRow)
            //            continue;
            //        if ("" + row.Cells[colType.Index].Value == "" + dataGridView.Rows[e.RowIndex].Cells[colType.Index].Value)
            //        {
            //            row.Cells[colAggregated.Index].Value = dataGridView.Rows[e.RowIndex].Cells[colAggregated.Index].Value;
            //            ValidateRow(row);
            //        }
            //    }
            //}
            //if (e.ColumnIndex == colType.Index)
            //{
            //    foreach (DataGridViewRow row in dataGridView.Rows)
            //    {
            //        if (row.Index == e.RowIndex || row.IsNewRow)
            //            continue;
            //        if ("" + row.Cells[colType.Index].Value == "" + dataGridView.Rows[e.RowIndex].Cells[colType.Index].Value)
            //        {
            //            dataGridView.Rows[e.RowIndex].Cells[colAggregated.Index].Value = row.Cells[colAggregated.Index].Value;
            //            ValidateRow(dataGridView.Rows[e.RowIndex]);
            //            break;
            //        }
            //    }
            //}
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("Periods");
            doc.AppendChild(root);

            bool valid = true;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.IsNewRow) continue;
                valid &= ValidateRow(row);

                XmlElement period = doc.CreateElement("Period");
                root.AppendChild(period);
                period.SetAttribute("Name", "" + row.Cells[colPeriod.Index].Value);
                period.SetAttribute("Type", "" + row.Cells[colType.Index].Value);
                period.SetAttribute("Sort", "" + row.Cells[colOrder.Index].Value);
                period.SetAttribute("Aggregated", "" + row.Cells[colWeight.Index].Value);
            
            }
            if (!valid)
            {
                FISCA.Presentation.Controls.MsgBox.Show("輸入資料有誤，請修正後再行儲存。", "內容錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return;
            }

            //string warningMsg = "變更節次名稱將會使得已使用該名稱之資料無法正確顯示於介面上，但並不會影響已儲存資料之正確性！\n是否儲存變更？";
            //if (FISCA.Presentation.Controls.MsgBox.Show(warningMsg, "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            //    return;

            DSXmlHelper helper = new DSXmlHelper("Lists");
            helper.AddElement("List");
            helper.AddElement("List", "Content", root.OuterXml, true);
            helper.AddElement("List", "Condition");
            helper.AddElement("List/Condition", "Name", Config.LIST_PERIODS_NAME);
            try
            {
                Config.Update(new DSRequest(helper));
            }
            catch (Exception exception)
            {
                FISCA.Presentation.Controls.MsgBox.Show("更新失敗 :" + exception.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                Config.Reset(Config.LIST_PERIODS);
                //FISCA.Presentation.Controls.MsgBox.Show("資料重設成功，新設定已生效。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("資料重設失敗，新設定值將於下次啟動系統後生效!", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            this.Close();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}