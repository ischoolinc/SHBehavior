using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using FISCA.Presentation.Controls;

namespace SHSchool.Behavior.ClassExtendControls
{
    public partial class DemeritNotificatForm : BaseForm
    {
        private DateTime _startDate;
        private DateTime _endDate;
        private bool _startTextBoxOK = false;
        private bool _endTextBoxOK = false;
        private bool _printable = false;

        private DisciplineNotificationPreference mDisciplineNotificationPreference;

        private DateTime StartDate
        {
            get { return _startDate; }
        }

        private DateTime EndDate
        {
            get { return _endDate; }
        }

        private DateTime GetMonthFirstDay(DateTime inputDate)
        {
            return DateTime.Parse(inputDate.Year + "/" + inputDate.Month + "/1");
        }

        private DateTime GetWeekFirstDay(DateTime inputDate)
        {
            switch (inputDate.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    return inputDate;
                case DayOfWeek.Tuesday:
                    return inputDate.AddDays(-1);
                case DayOfWeek.Wednesday:
                    return inputDate.AddDays(-2);
                case DayOfWeek.Thursday:
                    return inputDate.AddDays(-3);
                case DayOfWeek.Friday:
                    return inputDate.AddDays(-4);
                case DayOfWeek.Saturday:
                    return inputDate.AddDays(-5);
                default:
                    return inputDate.AddDays(-6);
            }
        }

        private void InitialDateRange()
        {
            switch (mDisciplineNotificationPreference.DateModeRangeMode)
            {
                case DateRangeMode.Month: //月
                    {
                        DateTime a = DateTime.Today;
                        a = GetMonthFirstDay(a.AddMonths(-1));
                        textBoxX1.Text = a.ToShortDateString();
                        textBoxX2.Text = a.AddMonths(1).AddDays(-1).ToShortDateString();
                        break;
                    }
                case DateRangeMode.Week: //週
                    {
                        DateTime b = DateTime.Today;
                        b = GetWeekFirstDay(b.AddDays(-7));
                        textBoxX1.Text = b.ToShortDateString();
                        textBoxX2.Text = b.AddDays(5).ToShortDateString();
                        break;
                    }
                case DateRangeMode.Custom: //自定
                    {
                        //textBoxX1.Text = DateTime.Today.ToShortDateString();
                        //textBoxX2.Text = textBoxX1.Text;
                        break;
                    }
                default:
                    throw new Exception("Date Range Mode Error.");
            }
        }

        public DemeritNotificatForm()
        {
            InitializeComponent();

            _startDate = DateTime.Today;
            _endDate = DateTime.Today;
            textBoxX1.Text = _startDate.ToShortDateString();
            textBoxX2.Text = _endDate.ToShortDateString();

            mDisciplineNotificationPreference = DisciplineNotificationPreference.GetInstance();

            //textBoxX2.Enabled = (mDisciplineNotificationPreference.DateModeRangeMode == DateRangeMode.Custom) ? true : false;

            InitialDateRange();

            textBoxX1.Text = DateTime.Today.ToShortDateString();
            textBoxX2.Text = textBoxX1.Text;
        }

        private void DemeritNotificatForm_Load(object sender, EventArgs e)
        {

        }

        private bool ValidateRange(string startDate, string endDate)
        {
            DateTime a, b;
            a = DateTime.Parse(startDate);
            b = DateTime.Parse(endDate);

            if (DateTime.Compare(b, a) < 0)
            {
                _printable = false;
                return false;
            }
            else
            {
                _printable = true;
                _startDate = a;
                _endDate = b;
                return true;
            }
        }

        protected bool Validate(string date)
        {
            DateTime a;
            return DateTime.TryParse(date, out a);
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (_printable == true)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }

            mDisciplineNotificationPreference.StartDate = textBoxX1.Text;
            mDisciplineNotificationPreference.EndDate = textBoxX2.Text;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void textBoxX1_TextChanged(object sender, EventArgs e)
        {
            if (Validate(textBoxX1.Text))
            {
                errorProvider1.Clear();
                _startTextBoxOK = true;
                if (_endTextBoxOK)
                {
                    if (!ValidateRange(textBoxX1.Text, textBoxX2.Text))
                        errorProvider1.SetError(textBoxX1, "日期區間錯誤");
                    else
                    {
                        errorProvider1.Clear();
                        errorProvider2.Clear();
                    }
                }
            }
            else
            {
                errorProvider1.SetError(textBoxX1, "日期格式錯誤");
                _startTextBoxOK = false;
                _printable = false;
            }
        }

        private void textBoxX2_TextChanged(object sender, EventArgs e)
        {
            if (Validate(textBoxX2.Text))
            {
                errorProvider2.Clear();
                _endTextBoxOK = true;
                if (_startTextBoxOK)
                {
                    if (!ValidateRange(textBoxX1.Text, textBoxX2.Text))
                        errorProvider2.SetError(textBoxX2, "日期區間錯誤");
                    else
                    {
                        errorProvider1.Clear();
                        errorProvider2.Clear();
                    }
                }
            }
            else
            {
                errorProvider2.SetError(textBoxX2, "日期格式錯誤");
                _endTextBoxOK = false;
                _printable = false;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DisciplineNotificationConfigForm configForm = new DisciplineNotificationConfigForm();

            if (configForm.ShowDialog() == DialogResult.OK)
                InitialDateRange();
        }
    }
}
