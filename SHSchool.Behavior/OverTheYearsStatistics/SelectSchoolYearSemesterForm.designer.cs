namespace SHSchool.Behavior
{
    partial class SelectSchoolYearSemesterForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該公開 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.numSchoolYear = new System.Windows.Forms.NumericUpDown();
            this.numSemester = new System.Windows.Forms.NumericUpDown();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.buttonX1 = new DevComponents.DotNetBar.ButtonX();
            this.checkBoxX1 = new DevComponents.DotNetBar.Controls.CheckBoxX();
            ((System.ComponentModel.ISupportInitialize)(this.numSchoolYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numSemester)).BeginInit();
            this.SuspendLayout();
            // 
            // numSchoolYear
            // 
            this.numSchoolYear.Location = new System.Drawing.Point(66, 9);
            this.numSchoolYear.Maximum = new decimal(new int[] {
            150,
            0,
            0,
            0});
            this.numSchoolYear.Minimum = new decimal(new int[] {
            90,
            0,
            0,
            0});
            this.numSchoolYear.Name = "numSchoolYear";
            this.numSchoolYear.Size = new System.Drawing.Size(45, 25);
            this.numSchoolYear.TabIndex = 0;
            this.numSchoolYear.Value = new decimal(new int[] {
            90,
            0,
            0,
            0});
            // 
            // numSemester
            // 
            this.numSemester.Location = new System.Drawing.Point(169, 9);
            this.numSemester.Maximum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.numSemester.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numSemester.Name = "numSemester";
            this.numSemester.Size = new System.Drawing.Size(45, 25);
            this.numSemester.TabIndex = 0;
            this.numSemester.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(6, 12);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(60, 21);
            this.labelX1.TabIndex = 1;
            this.labelX1.Text = "學年度：";
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(122, 12);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(47, 21);
            this.labelX2.TabIndex = 1;
            this.labelX2.Text = "學期：";
            // 
            // buttonX1
            // 
            this.buttonX1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX1.BackColor = System.Drawing.Color.Transparent;
            this.buttonX1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonX1.Location = new System.Drawing.Point(144, 43);
            this.buttonX1.Name = "buttonX1";
            this.buttonX1.Size = new System.Drawing.Size(70, 23);
            this.buttonX1.TabIndex = 2;
            this.buttonX1.Text = "確定";
            this.buttonX1.Click += new System.EventHandler(this.buttonX1_Click);
            // 
            // checkBoxX1
            // 
            this.checkBoxX1.AutoSize = true;
            this.checkBoxX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.checkBoxX1.BackgroundStyle.Class = "";
            this.checkBoxX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.checkBoxX1.Location = new System.Drawing.Point(5, 48);
            this.checkBoxX1.Name = "checkBoxX1";
            this.checkBoxX1.Size = new System.Drawing.Size(121, 21);
            this.checkBoxX1.TabIndex = 3;
            this.checkBoxX1.Text = "列印所有學年度";
            this.checkBoxX1.CheckedChanged += new System.EventHandler(this.checkBoxX1_CheckedChanged);
            // 
            // SelectSchoolYearSemesterForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(222, 73);
            this.Controls.Add(this.checkBoxX1);
            this.Controls.Add(this.buttonX1);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.numSemester);
            this.Controls.Add(this.numSchoolYear);
            this.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "SelectSchoolYearSemesterForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "選擇學年度學期";
            ((System.ComponentModel.ISupportInitialize)(this.numSchoolYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numSemester)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        protected System.Windows.Forms.NumericUpDown numSchoolYear;
        protected System.Windows.Forms.NumericUpDown numSemester;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        protected DevComponents.DotNetBar.ButtonX buttonX1;
        private DevComponents.DotNetBar.Controls.CheckBoxX checkBoxX1;

    }
}