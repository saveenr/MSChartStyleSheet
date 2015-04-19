using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MSCHART = System.Windows.Forms.DataVisualization.Charting;
namespace MSChartTestApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            var now = System.DateTime.Now;

            var dt = new System.Data.DataTable();
            dt.Columns.Add("Day", typeof (System.DateTime));
            dt.Columns.Add("Count", typeof(int));

            dt.Rows.Add(now.AddDays(-15), 103);
            dt.Rows.Add(now.AddDays(-14), 118);
            dt.Rows.Add(now.AddDays(-13), 143);
            dt.Rows.Add(now.AddDays(-12), 75);
            dt.Rows.Add(now.AddDays(-11), 82);
            dt.Rows.Add(now.AddDays(-10), 113);
            dt.Rows.Add(now.AddDays(-9), 102);
            dt.Rows.Add(now.AddDays(-8), 114);
            dt.Rows.Add(now.AddDays(-7), 85);
            dt.Rows.Add(now.AddDays(-6), 87);
            dt.Rows.Add(now.AddDays(-5), 63);
            dt.Rows.Add(now.AddDays(-4), 74);
            dt.Rows.Add(now.AddDays(-3), 73);
            dt.Rows.Add(now.AddDays(-2), 108);
            dt.Rows.Add(now.AddDays(-1), 114);
            dt.Rows.Add(now.AddDays(-0), 56);

            this.chart1.DataSource = dt;

            this.chart1.Series.Clear();
            var series = new MSCHART.Series("Series1");
            series.ChartType = MSCHART.SeriesChartType.Column;

            series.XValueMember = dt.Columns[0].ColumnName;
            series.YValueMembers = dt.Columns[1].ColumnName;

            this.chart1.Series.Add(series);

            var style = new MSChartStylesheet.Stylesheet();
            style.XAxisFormat = "m";

            style.Format(this.chart1);
        }
    }
}
