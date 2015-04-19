using System;
using MSCHART = System.Windows.Forms.DataVisualization.Charting;
using SD = System.Drawing;

namespace MSChartStylesheet
{
    public class ChartBuilder
    {
        public System.Data.DataTable DataTable { get; set; }
        public int? Width { get; set; }
        public int? Height { get; set; }
        public ChartType? ChartType { get; set; }
        public string Name { get; set; }
        public MSChartStylesheet.Stylesheet StyleSheet { get; set; }
        public string XColumn;
        public string YColumn;
        public string SeriesColumn;

        private MSCHART.Chart Build()
        {
            if (this.XColumn == null)
            {
                throw new System.ArgumentNullException("XColumn");
            }

            if (this.YColumn == null)
            {
                throw new System.ArgumentNullException("YColumn");
            }

            var chart = new MSCHART.Chart();
            if (this.Name != null)
            {
                chart.Name = this.Name;

                chart.Titles.Clear();
                var title = new MSCHART.Title(this.Name);
                chart.Titles.Add(title);
            }

            var area = new MSCHART.ChartArea();

            chart.ChartAreas.Add(area);

            if (this.SeriesColumn!=null)
            {
                chart.DataBindCrossTable(this.DataTable.Rows, this.SeriesColumn, this.XColumn, this.YColumn, "");
                var legend = new MSCHART.Legend("LEGEND");
                chart.Legends.Add(legend);
            }
            else
            {
                chart.DataSource = this.DataTable;
                var series = new MSCHART.Series("Series1");
                series.XValueMember = this.XColumn;
                series.YValueMembers = this.YColumn;
                chart.Series.Add(series);
            }


            var t = this.ChartType.HasValue ? this.ChartType.Value : MSChartStylesheet.ChartType.Column;
            foreach (var s in chart.Series)
            {

                if (t == MSChartStylesheet.ChartType.Column)
                {
                    s.ChartType = MSCHART.SeriesChartType.Column;
                }
                else if (t == MSChartStylesheet.ChartType.Bar)
                {
                    s.ChartType = MSCHART.SeriesChartType.Bar;
                }
                else if (t == MSChartStylesheet.ChartType.Pie)
                {
                    s.ChartType = MSCHART.SeriesChartType.Pie;
                }
                else if (t == MSChartStylesheet.ChartType.DoughNut)
                {
                    s.ChartType = MSCHART.SeriesChartType.Doughnut;
                }
                else if (t == MSChartStylesheet.ChartType.Line)
                {
                    s.ChartType = MSCHART.SeriesChartType.Line;
                }
            }



            var stylesheet = this.StyleSheet ?? new MSChartStylesheet.Stylesheet();
            stylesheet.RadialLabelStyle = MSChartStylesheet.RadialLabelStyle.Outside;
            stylesheet.Format(chart);

            chart.Width = this.Width.HasValue ? this.Width.Value : 600;
            chart.Height = this.Height.HasValue ? this.Height.Value : 400;


            return chart;
        }

        public void Save(string filename)
        {
            var ext = System.IO.Path.GetExtension(filename).ToLower();
            System.Drawing.Imaging.ImageFormat fmt = System.Drawing.Imaging.ImageFormat.Png;

            if (ext == ".png")
            {
                fmt = System.Drawing.Imaging.ImageFormat.Png;
            }
            else if (ext == ".jpg" || ext == ".jpeg")
            {
                fmt = System.Drawing.Imaging.ImageFormat.Jpeg;
            }
            else
            {
                throw new SystemException("format not supported");
            }

            var chart = this.Build();
            chart.SaveImage(filename, fmt);
        }

        public void Show()
        {
            string filename = System.IO.Path.Combine( System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName()) + ".png";
            this.Save(filename);
            System.Diagnostics.Process.Start(filename);
        }
    }
}