using SD=System.Drawing;

namespace MSChartStylesheet
{
    public static class Colors
    {
        // Black & White Scale
        public static SD.Color White = BuildGray(255);
        public static SD.Color GrayLight0 = BuildGray(243);
        public static SD.Color GrayLight1 = BuildGray(230);
        public static SD.Color GrayLight2 = BuildGray(210);
        public static SD.Color GrayMedium0 = BuildGray(150);
        public static SD.Color GrayMedium1 = BuildGray(130);
        public static SD.Color GrayMedium2 = BuildGray(100);
        public static SD.Color GrayDark0 = BuildGray(77);
        public static SD.Color GrayDark1 = BuildGray(51);
        public static SD.Color Black = SD.Color.FromArgb(255, 0, 0, 0);

        
        public static SD.Color BlueTitle = SD.Color.FromArgb(255, 68, 198, 234);
        public static SD.Color GrayTitle = SD.Color.FromArgb(255, 100,100,100);

        // Blues
        public static SD.Color BlueMedium0 = BuildColor(68, 198, 234);
        public static SD.Color BlueMedium1 = BuildColor(10, 175, 220);

        public static SD.Color BlueMLight0 = BuildColor(196, 227, 244);
        public static SD.Color BlueMLight1 = BuildColor(111, 206, 233);

        public static SD.Color XBlueMedium = BuildColor(34, 142, 165);
        public static SD.Color XBlueDark0 = BuildColor(13, 117, 144);

        // BlueGray
        public static SD.Color BlueGrayDark1 = BuildColor(4, 64, 93);
        public static SD.Color BlueGrayMedium = BuildColor(70, 108, 134);
        public static SD.Color BlueGrayLight = BuildColor(129, 167, 193);

        // GreenBlue
        public static SD.Color GreenBlueMedium = BuildColor(38,108,120);
        public static SD.Color GreenBlueDark = BuildColor(14,74,94);

        // Greens
        public static SD.Color GreenLight = BuildColor(111, 194, 118);
        public static SD.Color GreenLight1 = BuildColor(82, 188, 115);
        public static SD.Color GreenMedium = BuildColor(34, 171, 74);

        // Oranges
        public static SD.Color OrangeLight = BuildColor(242, 180, 96);
        public static SD.Color OrangeMedium = BuildColor(237, 160, 48);

        // Reds
        public static SD.Color RedMedium = BuildColor(239, 62, 54);
        public static SD.Color RedDark = BuildColor(191, 30, 45);

        // Wheat
        public static SD.Color WheatLight = BuildColor(252, 250, 206);
        public static SD.Color WheatDark = BuildColor(230, 184, 0);

        private static SD.Color BuildColor(int rgb)
        {
            return SD.Color.FromArgb((int)((uint)0xff000000 | (uint)rgb));
        }

        private static SD.Color BuildColor(int r,int g, int b)
        {
            return BuildColor((r << 16) | g << 8 | b);
        }

        private static SD.Color BuildGray(int n)
        {
            return BuildColor((n << 16) | n << 8 | n);
        }

        public static SD.Color BuildColorDelta(SD.Color srcolor, double hueDelta, double satDelta, double lumDelta)
        {
            double h, s, l;
            double nr, nb, ng;
            Normalize24BitRGB(srcolor.R, srcolor.G, srcolor.B, out nr, out ng, out nb);
            RGBToHSL(nr, ng, nb, out h, out s, out l);

            double nh = norm_hue(h + hueDelta);
            double ns = clip_01(s + satDelta);
            double nl = clip_01(l + lumDelta);

            HSLToRGB(nh, ns, nl, out nr, out ng, out nb);

            byte r, g, b;
            DeNormalize24BitRGB(nr, ng, nb, out r, out g, out b);
            return SD.Color.FromArgb(255, r, g, b);
        }

        public static void Normalize24BitRGB(byte r, byte g, byte b, out double R, out double G, out double B)
        {
            R = r / 255.0;
            G = g / 255.0;
            B = b / 255.0;
        }

        public static void DeNormalize24BitRGB(double R, double G, double B, out byte r, out byte g, out byte b)
        {
            r = (byte)(R * 255.0);
            g = (byte)(G * 255.0);
            b = (byte)(B * 255.0);
        }

        public static void RGBToHSL(double R, double G, double B, out double h, out double s, out double l)
        {
            double maxc = System.Math.Max(R, System.Math.Max(G, B));
            double minc = System.Math.Min(R, System.Math.Min(G, B));
            double delta = maxc - minc;

            l = (maxc + minc) / 2.0;

            // Handle case for r,g,b all have the same value
            if (maxc == minc)
            {
                // Black, White, or some shade of Gray -> No Chroma
                h = double.NaN;
                s = double.NaN;
                return;
            }

            // At this stage, we know R,G,B are not all set to the same value - i.e. there Chroma
            if (l < 0.5)
            {
                s = delta / (maxc + minc);
            }
            else
            {
                s = delta / (2.0 - maxc - minc);
            }

            double rc = (((maxc - R) / 6.0) + (delta / 2.0)) / delta;
            double gc = (((maxc - G) / 6.0) + (delta / 2.0)) / delta;
            double bc = (((maxc - B) / 6.0) + (delta / 2.0)) / delta;

            h = 0.0;

            if (R == maxc)
            {
                h = bc - gc;
            }
            else if (G == maxc)
            {
                h = (1.0 / 3.0) + rc - bc;
            }
            else if (B == maxc)
            {
                h = (2.0 / 3.0) + gc - rc;
            }

            h = norm_hue(h);
        }

        public static double clip_01(double v)
        {
            if (v < 0)
            {
                return 0;
            }
            if (v > 1)
            {
                return 1;
            }
            return v;
        }

        public static double norm_hue(double h)
        {
            if (h < 0)
            {
                return h + 1.0;
            }

            if (h > 1)
            {
                return h - 1.0;
            }

            return h;
        }

        public static void HSLToRGB(double H, double S, double L, out double r, out double g, out double b)
        {
            if (double.IsNaN(H) || S == 0) //HSL values = From 0 to 1
            {
                r = L; //RGB results = From 0 to 255
                g = L;
                b = L;
                return;
            }

            double m2 = (L < 0.5) ? L * (1.0 + S) : (L + S) - (S * L);
            double m1 = (2.0 * L) - m2;
            const double onethird = (1.0 / 3.0);
            r = 1.0 * hue_2_rgb(m1, m2, H + onethird);
            g = 1.0 * hue_2_rgb(m1, m2, H);
            b = 1.0 * hue_2_rgb(m1, m2, H - onethird);
        }

        private static double hue_2_rgb(double m1, double m2, double h)
        {
            h = norm_hue(h);

            if ((6.0 * h) < 1.0)
            {
                return (m1 + (m2 - m1) * 6.0 * h);
            }

            if ((2.0 * h) < 1.0)
            {
                return m2;
            }

            if ((3.0 * h) < 2.0)
            {
                return m1 + (m2 - m1) * ((2.0 / 3.0) - h) * 6.0;
            }

            return m1;
        }


    }
}