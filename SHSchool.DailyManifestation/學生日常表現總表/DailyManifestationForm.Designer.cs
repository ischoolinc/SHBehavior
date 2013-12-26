namespace SHSchool.DailyManifestation
{
    partial class DailyManifestationForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.lbHelp1 = new DevComponents.DotNetBar.LabelX();
            this.cbPrintMeritDemerit = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.cbAttendance = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.cbUpdateRecord = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.cbClearDemerit = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.cbSelectAllItem = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.cbDaily = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.groupPanel1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.groupPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.AutoSize = true;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Location = new System.Drawing.Point(234, 148);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 25);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 0;
            this.btnPrint.Text = "列印";
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(315, 148);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 1;
            this.btnExit.Text = "關閉";
            // 
            // lbHelp1
            // 
            this.lbHelp1.AutoSize = true;
            this.lbHelp1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbHelp1.BackgroundStyle.Class = "";
            this.lbHelp1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbHelp1.Location = new System.Drawing.Point(12, 12);
            this.lbHelp1.Name = "lbHelp1";
            this.lbHelp1.Size = new System.Drawing.Size(127, 21);
            this.lbHelp1.TabIndex = 2;
            this.lbHelp1.Text = "請選擇列印資料內容";
            // 
            // cbPrintMeritDemerit
            // 
            this.cbPrintMeritDemerit.AutoSize = true;
            this.cbPrintMeritDemerit.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.cbPrintMeritDemerit.BackgroundStyle.Class = "";
            this.cbPrintMeritDemerit.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.cbPrintMeritDemerit.Location = new System.Drawing.Point(21, 18);
            this.cbPrintMeritDemerit.Name = "cbPrintMeritDemerit";
            this.cbPrintMeritDemerit.Size = new System.Drawing.Size(86, 21);
            this.cbPrintMeritDemerit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbPrintMeritDemerit.TabIndex = 3;
            this.cbPrintMeritDemerit.Text = "獎勵/懲戒";
            // 
            // cbAttendance
            // 
            this.cbAttendance.AutoSize = true;
            this.cbAttendance.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.cbAttendance.BackgroundStyle.Class = "";
            this.cbAttendance.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.cbAttendance.Location = new System.Drawing.Point(272, 18);
            this.cbAttendance.Name = "cbAttendance";
            this.cbAttendance.Size = new System.Drawing.Size(80, 21);
            this.cbAttendance.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbAttendance.TabIndex = 4;
            this.cbAttendance.Text = "缺曠資料";
            // 
            // cbUpdateRecord
            // 
            this.cbUpdateRecord.AutoSize = true;
            this.cbUpdateRecord.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.cbUpdateRecord.BackgroundStyle.Class = "";
            this.cbUpdateRecord.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.cbUpdateRecord.Location = new System.Drawing.Point(136, 58);
            this.cbUpdateRecord.Name = "cbUpdateRecord";
            this.cbUpdateRecord.Size = new System.Drawing.Size(80, 21);
            this.cbUpdateRecord.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbUpdateRecord.TabIndex = 5;
            this.cbUpdateRecord.Text = "異動資料";
            // 
            // cbClearDemerit
            // 
            this.cbClearDemerit.AutoSize = true;
            this.cbClearDemerit.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.cbClearDemerit.BackgroundStyle.Class = "";
            this.cbClearDemerit.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.cbClearDemerit.Location = new System.Drawing.Point(136, 18);
            this.cbClearDemerit.Name = "cbClearDemerit";
            this.cbClearDemerit.Size = new System.Drawing.Size(107, 21);
            this.cbClearDemerit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbClearDemerit.TabIndex = 6;
            this.cbClearDemerit.Text = "懲戒銷過資料";
            // 
            // cbSelectAllItem
            // 
            this.cbSelectAllItem.AutoSize = true;
            this.cbSelectAllItem.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.cbSelectAllItem.BackgroundStyle.Class = "";
            this.cbSelectAllItem.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.cbSelectAllItem.Location = new System.Drawing.Point(36, 152);
            this.cbSelectAllItem.Name = "cbSelectAllItem";
            this.cbSelectAllItem.Size = new System.Drawing.Size(54, 21);
            this.cbSelectAllItem.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbSelectAllItem.TabIndex = 7;
            this.cbSelectAllItem.Text = "全選";
            // 
            // cbDaily
            // 
            this.cbDaily.AutoSize = true;
            this.cbDaily.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.cbDaily.BackgroundStyle.Class = "";
            this.cbDaily.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.cbDaily.Location = new System.Drawing.Point(21, 58);
            this.cbDaily.Name = "cbDaily";
            this.cbDaily.Size = new System.Drawing.Size(80, 21);
            this.cbDaily.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbDaily.TabIndex = 8;
            this.cbDaily.Text = "文字評量";
            // 
            // groupPanel1
            // 
            this.groupPanel1.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.cbPrintMeritDemerit);
            this.groupPanel1.Controls.Add(this.cbDaily);
            this.groupPanel1.Controls.Add(this.cbAttendance);
            this.groupPanel1.Controls.Add(this.cbUpdateRecord);
            this.groupPanel1.Controls.Add(this.cbClearDemerit);
            this.groupPanel1.Location = new System.Drawing.Point(12, 39);
            this.groupPanel1.Name = "groupPanel1";
            this.groupPanel1.Size = new System.Drawing.Size(378, 103);
            // 
            // 
            // 
            this.groupPanel1.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel1.Style.BackColorGradientAngle = 90;
            this.groupPanel1.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel1.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderBottomWidth = 1;
            this.groupPanel1.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel1.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderLeftWidth = 1;
            this.groupPanel1.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderRightWidth = 1;
            this.groupPanel1.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderTopWidth = 1;
            this.groupPanel1.Style.Class = "";
            this.groupPanel1.Style.CornerDiameter = 4;
            this.groupPanel1.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel1.Style.TextAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Center;
            this.groupPanel1.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel1.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.groupPanel1.StyleMouseDown.Class = "";
            this.groupPanel1.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.groupPanel1.StyleMouseOver.Class = "";
            this.groupPanel1.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.groupPanel1.TabIndex = 9;
            // 
            // DailyManifestationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(402, 185);
            this.Controls.Add(this.groupPanel1);
            this.Controls.Add(this.cbSelectAllItem);
            this.Controls.Add(this.lbHelp1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnPrint);
            this.Name = "DailyManifestationForm";
            this.Text = "學生日常表現總表";
            this.Load += new System.EventHandler(this.DailyManifestationForm_Load);
            this.groupPanel1.ResumeLayout(false);
            this.groupPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.LabelX lbHelp1;
        private DevComponents.DotNetBar.Controls.CheckBoxX cbPrintMeritDemerit;
        private DevComponents.DotNetBar.Controls.CheckBoxX cbAttendance;
        private DevComponents.DotNetBar.Controls.CheckBoxX cbUpdateRecord;
        private DevComponents.DotNetBar.Controls.CheckBoxX cbClearDemerit;
        private DevComponents.DotNetBar.Controls.CheckBoxX cbSelectAllItem;
        private DevComponents.DotNetBar.Controls.CheckBoxX cbDaily;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel1;
    }
}