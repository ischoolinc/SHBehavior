namespace SHSchool.DailyManifestation
{
    partial class NewSRoutineForm
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
               this.labelX1 = new DevComponents.DotNetBar.LabelX();
               this.btnClose = new DevComponents.DotNetBar.ButtonX();
               this.btnSave = new DevComponents.DotNetBar.ButtonX();
               this.linkLabel1 = new System.Windows.Forms.LinkLabel();
               this.cbSingeFile = new DevComponents.DotNetBar.Controls.CheckBoxX();
               this.checkBoxX1 = new DevComponents.DotNetBar.Controls.CheckBoxX();
               this.labelX2 = new DevComponents.DotNetBar.LabelX();
               this.SuspendLayout();
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
               this.labelX1.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
               this.labelX1.Location = new System.Drawing.Point(14, 18);
               this.labelX1.Name = "labelX1";
               this.labelX1.Size = new System.Drawing.Size(321, 21);
               this.labelX1.TabIndex = 5;
               this.labelX1.Text = "列印學生日常表現總表除產生報表,您可以有以下選擇:";
               // 
               // btnClose
               // 
               this.btnClose.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
               this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
               this.btnClose.AutoSize = true;
               this.btnClose.BackColor = System.Drawing.Color.Transparent;
               this.btnClose.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
               this.btnClose.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
               this.btnClose.Location = new System.Drawing.Point(314, 161);
               this.btnClose.Name = "btnClose";
               this.btnClose.Size = new System.Drawing.Size(75, 25);
               this.btnClose.TabIndex = 4;
               this.btnClose.Text = "關閉";
               this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
               // 
               // btnSave
               // 
               this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
               this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
               this.btnSave.AutoSize = true;
               this.btnSave.BackColor = System.Drawing.Color.Transparent;
               this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
               this.btnSave.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
               this.btnSave.Location = new System.Drawing.Point(233, 161);
               this.btnSave.Name = "btnSave";
               this.btnSave.Size = new System.Drawing.Size(75, 25);
               this.btnSave.TabIndex = 3;
               this.btnSave.Text = "列印";
               this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
               // 
               // linkLabel1
               // 
               this.linkLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
               this.linkLabel1.AutoSize = true;
               this.linkLabel1.BackColor = System.Drawing.Color.Transparent;
               this.linkLabel1.Location = new System.Drawing.Point(9, 168);
               this.linkLabel1.Name = "linkLabel1";
               this.linkLabel1.Size = new System.Drawing.Size(60, 17);
               this.linkLabel1.TabIndex = 6;
               this.linkLabel1.TabStop = true;
               this.linkLabel1.Text = "假別設定";
               this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
               // 
               // cbSingeFile
               // 
               this.cbSingeFile.AutoSize = true;
               this.cbSingeFile.BackColor = System.Drawing.Color.Transparent;
               // 
               // 
               // 
               this.cbSingeFile.BackgroundStyle.Class = "";
               this.cbSingeFile.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
               this.cbSingeFile.Location = new System.Drawing.Point(50, 89);
               this.cbSingeFile.Name = "cbSingeFile";
               this.cbSingeFile.Size = new System.Drawing.Size(276, 21);
               this.cbSingeFile.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
               this.cbSingeFile.TabIndex = 7;
               this.cbSingeFile.Text = "進行單檔列印(一名學生一張報表另存新檔)";
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
               this.checkBoxX1.Location = new System.Drawing.Point(50, 55);
               this.checkBoxX1.Name = "checkBoxX1";
               this.checkBoxX1.Size = new System.Drawing.Size(147, 21);
               this.checkBoxX1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
               this.checkBoxX1.TabIndex = 9;
               this.checkBoxX1.Text = "上傳為學生電子報表";
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
               this.labelX2.Location = new System.Drawing.Point(69, 120);
               this.labelX2.Name = "labelX2";
               this.labelX2.Size = new System.Drawing.Size(257, 21);
               this.labelX2.TabIndex = 13;
               this.labelX2.Text = "(檔名規格:學號_身分證號_班級_座號_姓名)";
               // 
               // NewSRoutineForm
               // 
               this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
               this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
               this.ClientSize = new System.Drawing.Size(398, 194);
               this.Controls.Add(this.labelX2);
               this.Controls.Add(this.checkBoxX1);
               this.Controls.Add(this.cbSingeFile);
               this.Controls.Add(this.linkLabel1);
               this.Controls.Add(this.labelX1);
               this.Controls.Add(this.btnClose);
               this.Controls.Add(this.btnSave);
               this.DoubleBuffered = true;
               this.Name = "NewSRoutineForm";
               this.Text = "學生日常表現總表";
               this.Load += new System.EventHandler(this.NewSRoutineForm_Load);
               this.ResumeLayout(false);
               this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.ButtonX btnClose;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private DevComponents.DotNetBar.Controls.CheckBoxX cbSingeFile;
        private DevComponents.DotNetBar.Controls.CheckBoxX checkBoxX1;
        private DevComponents.DotNetBar.LabelX labelX2;
    }
}