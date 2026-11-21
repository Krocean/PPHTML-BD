// Compares two pictures of the same page (the modern browser's and the IE11 engine's) so that the build can say
// how much an automatic conversion changed the look, and writes them side by side for a human or model to judge.
// C# 5 only (compiled by Windows PowerShell's Add-Type). Reference: System.Drawing.
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace ImgKit
{
    public static class Compare
    {
        static byte[] Small(Bitmap src, int w, int h)
        {
            using (var b = new Bitmap(w, h, PixelFormat.Format24bppRgb))
            {
                using (var g = Graphics.FromImage(b)) { g.InterpolationMode = InterpolationMode.HighQualityBilinear; g.PixelOffsetMode = PixelOffsetMode.HighQuality; g.DrawImage(src, new Rectangle(0, 0, w, h)); }
                var data = b.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                var buf = new byte[data.Stride * h];
                Marshal.Copy(data.Scan0, buf, 0, buf.Length);
                b.UnlockBits(data);
                var outp = new byte[w * h * 3];
                for (int y = 0; y < h; y++) Buffer.BlockCopy(buf, y * data.Stride, outp, y * w * 3, w * 3);
                return outp;
            }
        }

        // A small JPEG of a picture, as base64 (for the slide picker inside the pages).
        public static string Thumb(string file, int width, long quality)
        {
            using (var src = Image.FromFile(file))
            {
                int h = (int)Math.Round(width * (double)src.Height / src.Width);
                using (var b = new Bitmap(width, h, PixelFormat.Format24bppRgb))
                {
                    using (var g = Graphics.FromImage(b)) { g.InterpolationMode = InterpolationMode.HighQualityBicubic; g.PixelOffsetMode = PixelOffsetMode.HighQuality; g.DrawImage(src, new Rectangle(0, 0, width, h)); }
                    ImageCodecInfo enc = null;
                    foreach (var c in ImageCodecInfo.GetImageEncoders()) if (c.MimeType == "image/jpeg") enc = c;
                    using (var ms = new System.IO.MemoryStream())
                    using (var prm = new EncoderParameters(1))
                    {
                        prm.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
                        b.Save(ms, enc, prm);
                        return Convert.ToBase64String(ms.ToArray());
                    }
                }
            }
        }

        // All pictures on one sheet (numbered, in order). cellWidth: width of one picture on the sheet.
        // The main colours of a picture: "RRGGBB:share;..." (share in per cent of the picture), most used first.
        public static string Palette(string file, int count)
        {
            const int N = 96;
            using (var src = Image.FromFile(file))
            using (var small = new Bitmap(N, N, PixelFormat.Format24bppRgb))
            {
                using (var g = Graphics.FromImage(small)) { g.InterpolationMode = InterpolationMode.HighQualityBilinear; g.DrawImage(src, 0, 0, N, N); }
                var bins = new System.Collections.Generic.Dictionary<int, int[]>();
                for (int y = 0; y < N; y++)
                    for (int x = 0; x < N; x++)
                    {
                        Color c = small.GetPixel(x, y);
                        int key = ((c.R >> 4) << 8) | ((c.G >> 4) << 4) | (c.B >> 4);
                        int[] b;
                        if (!bins.TryGetValue(key, out b)) { b = new int[4]; bins[key] = b; }
                        b[0]++; b[1] += c.R; b[2] += c.G; b[3] += c.B;
                    }
                var list = new System.Collections.Generic.List<int[]>(bins.Values);
                list.Sort(delegate(int[] p, int[] q) { return q[0].CompareTo(p[0]); });
                var merged = new System.Collections.Generic.List<int[]>();
                foreach (var b in list)
                {
                    int r = b[1] / b[0], gg = b[2] / b[0], bb = b[3] / b[0];
                    bool done = false;
                    foreach (var o in merged)
                    {
                        int dr = o[1] - r, dg = o[2] - gg, db = o[3] - bb;
                        if (dr * dr + dg * dg + db * db < 34 * 34) { o[0] += b[0]; done = true; break; }
                    }
                    if (!done) merged.Add(new int[] { b[0], r, gg, bb });
                }
                merged.Sort(delegate(int[] p, int[] q) { return q[0].CompareTo(p[0]); });
                var sb = new System.Text.StringBuilder();
                for (int i = 0; i < merged.Count && i < count; i++)
                {
                    var o = merged[i];
                    if (sb.Length > 0) sb.Append(';');
                    sb.Append(o[1].ToString("X2")).Append(o[2].ToString("X2")).Append(o[3].ToString("X2")).Append(':');
                    sb.Append((o[0] * 100.0 / (N * N)).ToString("0.#", System.Globalization.CultureInfo.InvariantCulture));
                }
                return sb.ToString();
            }
        }

        public static string Sheet(string[] files, string outFile, int columns, int cellWidth) { return Sheet(files, outFile, columns, cellWidth, null); }

        // labels: the text above each picture (default: "page N")
        public static string Sheet(string[] files, string outFile, int columns, int cellWidth, string[] labels)
        {
            if (files == null || files.Length == 0) return "";
            int cols = Math.Max(1, Math.Min(columns, files.Length)), rows = (files.Length + cols - 1) / cols, cellH = 0, pad = 14, top = 26;
            using (var first = Image.FromFile(files[0])) cellH = (int)Math.Round(cellWidth * (double)first.Height / first.Width);
            using (var sheet = new Bitmap(cols * (cellWidth + pad) + pad, rows * (cellH + top + pad) + pad, PixelFormat.Format24bppRgb))
            {
                using (var g = Graphics.FromImage(sheet))
                using (var f = new Font("Segoe UI", 14f, FontStyle.Bold, GraphicsUnit.Pixel))
                {
                    g.Clear(Color.FromArgb(40, 40, 40));
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    for (int i = 0; i < files.Length; i++)
                    {
                        int x = pad + (i % cols) * (cellWidth + pad), y = pad + (i / cols) * (cellH + top + pad);
                        g.DrawString(labels != null && i < labels.Length && !string.IsNullOrEmpty(labels[i]) ? labels[i] : "page " + (i + 1), f, Brushes.White, x, y + 3);
                        try { using (var img = Image.FromFile(files[i])) g.DrawImage(img, new Rectangle(x, y + top, cellWidth, cellH)); } catch { }
                    }
                }
                sheet.Save(outFile, ImageFormat.Png);
            }
            return outFile;
        }

        // Returns "similarity|differingCells|totalCells|worstCellColumn|worstCellRow" (similarity 0-100; cells of a 16 x 9 grid).
        // sideBySide: path of a picture with both versions next to each other (left = a, right = b), or "".
        public static string Run(string a, string b, string sideBySide, int maxWidth)
        {
            using (var ia = (Bitmap)Image.FromFile(a))
            using (var ib = (Bitmap)Image.FromFile(b))
            {
                const int W = 320, H = 180, CX = 16, CY = 9;
                byte[] pa = Small(ia, W, H), pb = Small(ib, W, H);
                long total = 0; var cell = new long[CX * CY];
                for (int y = 0; y < H; y++)
                    for (int x = 0; x < W; x++)
                    {
                        int o = (y * W + x) * 3;
                        int d = Math.Abs(pa[o] - pb[o]) + Math.Abs(pa[o + 1] - pb[o + 1]) + Math.Abs(pa[o + 2] - pb[o + 2]);
                        total += d; cell[(y * CY / H) * CX + (x * CX / W)] += d;
                    }
                double sim = 100.0 - (total / (double)(W * H * 3 * 255)) * 100.0;
                int per = (W / CX) * (H / CY) * 3, differing = 0, worst = 0;
                for (int i = 0; i < cell.Length; i++)
                {
                    if (cell[i] / (double)per > 12.0) differing++;       // mean difference per channel above 12 of 255
                    if (cell[i] > cell[worst]) worst = i;
                }
                if (!string.IsNullOrEmpty(sideBySide))
                {
                    int half = Math.Min(maxWidth / 2, Math.Max(ia.Width, ib.Width));
                    int hh = (int)Math.Round(half * (double)ia.Height / ia.Width);
                    using (var outp = new Bitmap(half * 2 + 12, hh + 34, PixelFormat.Format24bppRgb))
                    {
                        using (var g = Graphics.FromImage(outp))
                        {
                            g.Clear(Color.FromArgb(40, 40, 40));
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(ia, new Rectangle(0, 34, half, hh));
                            g.DrawImage(ib, new Rectangle(half + 12, 34, half, hh));
                            using (var f = new Font("Segoe UI", 13f, FontStyle.Bold, GraphicsUnit.Pixel))
                            {
                                g.DrawString("modern browser (what the page was written for)", f, Brushes.White, 8, 8);
                                g.DrawString("slide show engine (IE11), after conversion", f, Brushes.White, half + 20, 8);
                            }
                        }
                        outp.Save(sideBySide, ImageFormat.Png);
                    }
                }
                return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:F1}|{1}|{2}|{3}|{4}", sim, differing, CX * CY, worst % CX + 1, worst / CX + 1);
            }
        }
    }
}
