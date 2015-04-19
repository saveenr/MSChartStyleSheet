using System;
using MSCHART = System.Windows.Forms.DataVisualization.Charting;
using SD = System.Drawing;

namespace MSChartStylesheet
{
    public class Stylesheet : IDisposable
    {
        public bool UseGradients = false;

        // Points
        public double? PointWidth;
        public SD.Font PointLabelFont;
        public SD.Color PointLabelColor = Colors.GrayMedium2;
        public float PointLabelFontSize = 8.0f;

        // Title
        public SD.Font TitleFont;
        public float TitleFontSize = 30.0f;
        public SD.Color TitleColor = Colors.BlueTitle;
        public SD.ContentAlignment TitleAlignment = SD.ContentAlignment.TopLeft;

        // Entire Chart
        public SD.Color BackgroundColor = Colors.White;

        // Axis & Grid
        public bool AxisXMajorGridEnabled = true;
        public bool AxisYMajorGridEnabled = true;
        public SD.Font AxisLabelFont;
        public SD.Font AxisTitleFont;
        public bool AxisMinorTickMark = false;
        public bool AxisMajorTickMark = false;
        public float AxisLabelFontSize = 8.0f;
        public float AxisTitleFontSize = 8.0f;

        // Legend
        public MSCHART.Docking LegendDocking = MSCHART.Docking.Top;
        public SD.Font LegendFont;
        public SD.Font LegendTitleFont;
        public SD.StringAlignment LegendTitleAlignment = SD.StringAlignment.Near;
        public float LegendFontSize = 8.0f;
        public float LegendTitleFontSize = 8.0f;

        // General
        public SD.Color TextColor = Colors.GrayMedium2;
        public SD.Color LineColor = Colors.GrayLight1;

        // Line
        public SD.Color MarkerColor = Colors.White;
        public SD.Color MarkerBorderColor = Colors.Black;
        public MSCHART.MarkerStyle LineMarker = MSCHART.MarkerStyle.Circle;

        // Pie 
        public RadialLabelStyle RadialLabelStyle = RadialLabelStyle.Inside;
        public SD.Color RadialLabelLineColor = Colors.GrayLight2;
        public int RadialLabelLineSize = 1;

        // Palette
        public System.Drawing.Color[] Palette = new[]
            {
                Colors.BlueGrayMedium, Colors.BlueGrayDark1, Colors.GrayLight2, Colors.GreenMedium,
                Colors.GreenBlueMedium, Colors.GreenBlueDark,
                Colors.WheatLight, Colors.WheatDark, Colors.OrangeLight, Colors.OrangeMedium
            };

        // Value Formatting
        // http://msdn.microsoft.com/en-us/library/az4se3k1.aspx
        public string XAxisFormat;

        // Fonts
        public string FontName = "Segoe UI";
        public string TitleFontName = "Segoe UI Light";

        public Stylesheet()
        {
        }

        private void InitFonts()
        {
            if (this.PointLabelFont == null)
            {
                this.PointLabelFont = new SD.Font(this.FontName, this.PointLabelFontSize);
                this.LegendFont = new SD.Font(this.FontName, this.LegendFontSize);
                this.LegendTitleFont = new SD.Font(this.FontName, this.LegendTitleFontSize);
                this.AxisLabelFont = new SD.Font(this.FontName, this.AxisLabelFontSize);
                this.AxisTitleFont = new SD.Font(this.FontName, this.AxisTitleFontSize);
                this.TitleFont = new SD.Font(this.TitleFontName, this.TitleFontSize);
            }
        }

        public void Format(MSCHART.Chart chart)
        {
            this.InitFonts();
            FormatPalette(chart);
            FormatSeries(chart);
            FormatAxis(chart);
            FormatLegend(chart);
            FormatTitles(chart);
            FormatBackground(chart);
            FormatRendering(chart);
        }

        private void FormatBackground(MSCHART.Chart chart)
        {
            chart.BackColor = BackgroundColor;
            foreach (var area in chart.ChartAreas)
            {
                area.BackColor = this.BackgroundColor;
            }
        }

        private void FormatPalette(MSCHART.Chart chart)
        {
            chart.Palette = MSCHART.ChartColorPalette.None;
            chart.PaletteCustomColors = this.Palette;
        }

        private void FormatRendering(MSCHART.Chart chart)
        {
            // Pick some options for high quality rendering
            chart.AntiAliasing = MSCHART.AntiAliasingStyles.All;
            chart.TextAntiAliasingQuality = MSCHART.TextAntiAliasingQuality.High;
            chart.IsSoftShadows = true;
        }

        private void FormatSeries(MSCHART.Chart chart)
        {
            for (int i = 0; i < chart.Series.Count; i++)
            {
                var c = chart.PaletteCustomColors[i%chart.PaletteCustomColors.Length];
                var series = chart.Series[i];

                if (PointWidth.HasValue)
                {
                    series["PointWidth"] = PointWidth.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                // Format the series labels
                series.Font = this.PointLabelFont;
                series.LabelForeColor = this.PointLabelColor;

                if (series.ChartType == MSCHART.SeriesChartType.Line)
                {
                    series.MarkerStyle = this.LineMarker;
                    series.MarkerColor = this.MarkerColor;
                    series.MarkerBorderColor = c;
                }
                else if (series.ChartType == MSCHART.SeriesChartType.Pie ||
                         series.ChartType == MSCHART.SeriesChartType.Doughnut)
                {
                    series["PieLabelStyle"] = this.RadialLabelStyle.ToString();
                    series["PieLineColor"] = SD.ColorTranslator.ToHtml(this.RadialLabelLineColor);
                    series["LabelsRadialLineSize"] = this.RadialLabelLineSize.ToString();
                }
            }

            if (this.UseGradients)
            {
                double hueDelta = -0.05;
                double satDelta = +0.15;
                double lumDelta = -0.2;

                for (int i = 0; i < chart.Series.Count; i++)
                {
                    var c = chart.PaletteCustomColors[i%chart.PaletteCustomColors.Length];
                    var series = chart.Series[i];
                    series.Color = c;
                    series.BackSecondaryColor = MSChartStylesheet.Colors.BuildColorDelta(c, hueDelta, satDelta, lumDelta);
                    if (series.ChartType == MSCHART.SeriesChartType.Pie)
                    {
                        series.BackGradientStyle = MSCHART.GradientStyle.Center;
                    }
                    else
                    {
                        series.BackGradientStyle = MSCHART.GradientStyle.TopBottom;
                    }
                }

            }
        }

        private void FormatAxis(MSCHART.Chart chart)
        {
            foreach (var area in chart.ChartAreas)
            {

                foreach (var axis in area.Axes)
                {
                    axis.LabelStyle.Font = this.AxisLabelFont;
                    axis.LabelStyle.ForeColor = this.TextColor;
                    axis.LineColor = this.LineColor;
                    axis.MinorTickMark.Enabled = this.AxisMajorTickMark;
                    axis.MajorTickMark.Enabled = this.AxisMinorTickMark;
                    axis.MajorGrid.LineColor = this.LineColor;
                    axis.MinorGrid.LineColor = this.LineColor;
                }

                // Show all Categories on X Axis
                area.AxisX.Interval = 1;

                // Set Formatting for Category Axis
                area.AxisX.LabelStyle.Format = this.XAxisFormat;

                // Set visibility of grid lines
                area.AxisX.MajorGrid.Enabled = AxisXMajorGridEnabled;
                area.AxisY.MajorGrid.Enabled = AxisYMajorGridEnabled;

                foreach (var axis in area.Axes)
                {
                    axis.TitleFont = this.AxisTitleFont;
                    axis.TitleForeColor = this.TextColor;
                }
            }
        }

        private void FormatLegend(MSCHART.Chart chart)
        {
            foreach (var legend in chart.Legends)
            {
                legend.Font = this.LegendFont;
                legend.ForeColor = this.TextColor;
                legend.TitleAlignment = this.LegendTitleAlignment;
                legend.TitleFont = LegendTitleFont;
                legend.TitleForeColor = this.TextColor;
                legend.Docking = this.LegendDocking;
            }
        }

        public void FormatTitles(MSCHART.Chart chart)
        {
            foreach (var title in chart.Titles)
            {
                title.Font = this.TitleFont;
                title.ForeColor = this.TitleColor;
                title.Alignment = this.TitleAlignment;
            }
        }

        public void Dispose()
        {
            if (this.AxisLabelFont!=null)
            {
                this.AxisLabelFont.Dispose();
            }

            if (this.AxisTitleFont!=null)
            {
                this.AxisLabelFont.Dispose();
            }

            if (this.LegendFont!=null)
            {
                this.LegendFont.Dispose();
            }

            if (this.LegendTitleFont!=null)
            {
                this.LegendTitleFont.Dispose();
            }

            if (this.PointLabelFont!=null)
            {
                this.LegendTitleFont.Dispose();
            }

            if (this.TitleFont!=null)
            {
                this.TitleFont.Dispose();
            }
        }
    }
}