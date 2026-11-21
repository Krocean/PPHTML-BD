using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

// Prepares an HTML page for the browser control that PowerPoint hosts (the Internet Explorer 11 engine):
//  - one self-contained document (local style sheets, scripts, images and fonts are pulled in),
//  - CSS custom properties resolved when that can be done without changing the meaning,
//  - the small script blocks that connect the page to its host ("harness") added,
//  - a list of constructs the IE11 engine does not support.
// C# 5 only (compiled by Windows PowerShell's Add-Type).
namespace PptHtml
{
    public class Issue
    {
        public string Severity;   // "error" = will not work in IE11, "warn" = may behave differently
        public string Rule;
        public string Message;
        public string Snippet;
        public int Line;
        public string Lines = "";   // all line numbers where the rule fired in this block
        public int Count = 1;
    }

    public class PrepResult
    {
        public string Html = "";
        public List<Issue> Issues = new List<Issue>();
        public List<string> Notes = new List<string>();
        public List<string> External = new List<string>();
        public List<string> Missing = new List<string>();
        public bool UsesViewportUnits;
        public long InlinedBytes;
    }

    // A page between the steps of its preparation (Begin -> ApplyConversion -> Finish).
    public class PrepState
    {
        internal List<Seg> Segs;
        internal List<Seg> ScriptSegs = new List<Seg>(), StyleSegs = new List<Seg>();
        internal string Deferred = "";
        public PrepResult R = new PrepResult();
        public string[] Scripts = new string[0];        // inline JavaScript blocks, in page order
        public bool[] ScriptIsModule = new bool[0];
        public string[] Styles = new string[0];         // style sheets, in page order (media attributes written as @media)
        public string[] InlineStyles = new string[0];   // values of style attributes, in page order
        public string Doc = "";                         // the page with its files pulled in, before any conversion
        public bool NeedsJs, NeedsCss, OptOut;          // the rules found script syntax / CSS that IE11 lacks
        public bool Converted, ScriptsConverted;
        public string TableJson = "";
        public string[] Allow = new string[0];
        public string ExtraCss = "";                    // added in a style block of its own (embedded fonts)
    }

    class Seg
    {
        public int Kind;          // 0 markup, 1 comment, 2 script, 3 style
        public string Text;       // whole text (markup / comment) or inner content (script / style)
        public string Open;       // open tag for script / style
        public string Close;      // close tag for script / style
    }

    public static class HtmlPrep
    {
        const RegexOptions IC = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;
        public const string HeadOpen = "<!--ih:head-->", HeadClose = "<!--/ih:head-->";
        public const string TailOpen = "<!--ih:tail-->", TailClose = "<!--/ih:tail-->";
        const string XUa = "<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\"><meta charset=\"utf-8\">";

        // ------------------------------------------------------------------ segments

        static readonly Regex SegRx = new Regex(@"<!--[\s\S]*?-->|<script\b[^>]*>[\s\S]*?</script\s*>|<style\b[^>]*>[\s\S]*?</style\s*>", IC);

        static List<Seg> Split(string html)
        {
            var list = new List<Seg>();
            int pos = 0;
            foreach (Match m in SegRx.Matches(html))
            {
                if (m.Index > pos) list.Add(new Seg { Kind = 0, Text = html.Substring(pos, m.Index - pos) });
                string t = m.Value;
                if (t.StartsWith("<!--")) list.Add(new Seg { Kind = 1, Text = t });
                else
                {
                    int gt = t.IndexOf('>');
                    int lt = t.LastIndexOf("</", StringComparison.Ordinal);
                    bool script = t.Length > 2 && (t[1] == 's' || t[1] == 'S') && (t[2] == 'c' || t[2] == 'C');
                    list.Add(new Seg { Kind = script ? 2 : 3, Open = t.Substring(0, gt + 1), Text = t.Substring(gt + 1, lt - gt - 1), Close = t.Substring(lt) });
                }
                pos = m.Index + m.Length;
            }
            if (pos < html.Length) list.Add(new Seg { Kind = 0, Text = html.Substring(pos) });
            return list;
        }

        static string Join(List<Seg> segs)
        {
            var sb = new StringBuilder();
            foreach (var s in segs)
            {
                if (s.Kind >= 2) { sb.Append(s.Open); sb.Append(s.Text); sb.Append(s.Close); }
                else sb.Append(s.Text);
            }
            return sb.ToString();
        }

        static string Attr(string tag, string name)
        {
            Match m = Regex.Match(tag, @"\s" + name + @"\s*=\s*(?:""([^""]*)""|'([^']*)'|([^\s>]+))", IC);
            if (!m.Success) return null;
            if (m.Groups[1].Success) return m.Groups[1].Value;
            if (m.Groups[2].Success) return m.Groups[2].Value;
            return m.Groups[3].Value;
        }

        // ------------------------------------------------------------------ local resources

        static bool IsExternal(string url) { return Regex.IsMatch(url, @"^\s*(https?:)?//", IC); }
        static bool IsInert(string url) { return Regex.IsMatch(url, @"^\s*(data:|#|javascript:|mailto:|tel:|about:|blob:|\{\{)", IC) || url.Trim().Length == 0; }

        static string LocalPath(string url, string baseDir)
        {
            string u = url.Trim();
            if (u.StartsWith("file:", StringComparison.OrdinalIgnoreCase)) { u = Regex.Replace(u, @"^file:/+", "", IC); }
            int q = u.IndexOfAny(new[] { '?', '#' });
            if (q >= 0) u = u.Substring(0, q);
            try { u = Uri.UnescapeDataString(u); } catch { }
            u = u.Replace('/', '\\');
            try { return Path.IsPathRooted(u) ? Path.GetFullPath(u) : Path.GetFullPath(Path.Combine(baseDir, u)); }
            catch { return null; }
        }

        static string Mime(string path)
        {
            switch (Path.GetExtension(path).ToLowerInvariant())
            {
                case ".png": return "image/png";
                case ".jpg": case ".jpeg": return "image/jpeg";
                case ".gif": return "image/gif";
                case ".svg": return "image/svg+xml";
                case ".bmp": return "image/bmp";
                case ".ico": return "image/x-icon";
                case ".webp": return "image/webp";
                case ".woff": return "application/font-woff";
                case ".woff2": return "font/woff2";
                case ".ttf": return "application/x-font-ttf";
                case ".otf": return "application/x-font-opentype";
                case ".eot": return "application/vnd.ms-fontobject";
                case ".mp4": return "video/mp4";
                case ".mp3": return "audio/mpeg";
                case ".wav": return "audio/wav";
                default: return null;
            }
        }

        static string ReadText(string path)
        {
            byte[] b = File.ReadAllBytes(path);
            if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) return Encoding.Unicode.GetString(b, 2, b.Length - 2);
            if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return Encoding.UTF8.GetString(b, 3, b.Length - 3);
            return Encoding.UTF8.GetString(b);
        }

        // Decodes a page file: UTF-16 / UTF-8 by byte order mark, otherwise the charset the page declares, otherwise UTF-8.
        public static string ReadPage(string path)
        {
            byte[] b = File.ReadAllBytes(path);
            if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) return Encoding.Unicode.GetString(b, 2, b.Length - 2);
            if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(b, 2, b.Length - 2);
            if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return Encoding.UTF8.GetString(b, 3, b.Length - 3);
            string head = Encoding.ASCII.GetString(b, 0, Math.Min(b.Length, 4096));
            Match m = Regex.Match(head, @"<meta[^>]+charset\s*=\s*[""']?\s*([\w-]+)", IC);
            if (m.Success)
            {
                string cs = m.Groups[1].Value.ToLowerInvariant();
                if (cs != "utf-8" && cs != "utf8")
                {
                    try { return Encoding.GetEncoding(cs).GetString(b); } catch { }
                }
            }
            return Encoding.UTF8.GetString(b);
        }

        static string DataUri(string url, string baseDir, PrepResult r, string what)
        {
            if (IsInert(url)) return null;
            if (IsExternal(url)) { if (!r.External.Contains(url.Trim())) r.External.Add(url.Trim()); return null; }
            string p = LocalPath(url, baseDir);
            if (p == null || !File.Exists(p)) { string miss = url.Trim() + " (" + what + ")"; if (!r.Missing.Contains(miss)) r.Missing.Add(miss); return null; }
            string mime = Mime(p);
            if (mime == null) return null;
            if (mime == "image/webp") r.Issues.Add(new Issue { Severity = "error", Rule = "img-webp", Message = "WebP images do not display in IE11; convert the file to PNG or JPEG.", Snippet = url.Trim() });
            if (mime == "font/woff2") r.Issues.Add(new Issue { Severity = "warn", Rule = "font-woff2", Message = "WOFF2 fonts are ignored by IE11; add a WOFF source or rely on an installed font.", Snippet = url.Trim() });
            byte[] b = File.ReadAllBytes(p);
            b = Slim(b, ref mime, r, Path.GetFileName(p));
            r.InlinedBytes += b.Length;
            return "data:" + mime + ";base64," + Convert.ToBase64String(b);
        }

        // ------------------------------------------------------------------ pictures
        // A picture straight from a camera or a design tool is often several times larger than a slide can show.
        // Large files are scaled down to what a 4K screen can use, photographs stored as PNG or BMP become JPEG,
        // and a JPEG that carries a rotation note is turned upright (the slide show engine ignores the note).
        public static int ImageMaxSide = 3840;          // longest side kept, in px; 0 = leave every picture as it is
        public static int ImageMinBytes = 300 * 1024;   // smaller files are left alone (unless they need turning)

        static string Kb(long n) { return n >= 1024 * 1024 ? (n / 1048576.0).ToString("0.0", System.Globalization.CultureInfo.InvariantCulture) + " MB" : (n / 1024) + " KB"; }

        static bool Opaque(System.Drawing.Bitmap bmp)
        {
            var d = bmp.LockBits(new System.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height), System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            try
            {
                int n = Math.Abs(d.Stride) * d.Height;
                byte[] px = new byte[n];
                System.Runtime.InteropServices.Marshal.Copy(d.Scan0, px, 0, n);
                for (int i = 3; i < n; i += 4) if (px[i] < 250) return false;
                return true;
            }
            finally { bmp.UnlockBits(d); }
        }

        static byte[] Encode(System.Drawing.Bitmap bmp, bool jpeg, long quality)
        {
            using (var ms = new MemoryStream())
            {
                if (!jpeg) { bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png); return ms.ToArray(); }
                System.Drawing.Imaging.ImageCodecInfo codec = null;
                foreach (var c in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()) if (c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid) codec = c;
                using (var ep = new System.Drawing.Imaging.EncoderParameters(1))
                {
                    ep.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
                    using (var flat = new System.Drawing.Bitmap(bmp.Width, bmp.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb))
                    {
                        using (var g = System.Drawing.Graphics.FromImage(flat)) { g.Clear(System.Drawing.Color.White); g.DrawImageUnscaled(bmp, 0, 0); }
                        flat.Save(ms, codec, ep);
                    }
                }
                return ms.ToArray();
            }
        }

        static byte[] Slim(byte[] b, ref string mime, PrepResult r, string name)
        {
            if (ImageMaxSide <= 0) return b;
            if (mime != "image/png" && mime != "image/jpeg" && mime != "image/bmp") return b;
            try
            {
                using (var ms = new MemoryStream(b))
                using (var img = System.Drawing.Image.FromStream(ms))
                {
                    int w = img.Width, h = img.Height, turn = 1;
                    if (mime == "image/jpeg") { try { if (Array.IndexOf(img.PropertyIdList, 0x0112) >= 0) turn = BitConverter.ToUInt16(img.GetPropertyItem(0x0112).Value, 0); } catch { } }
                    bool rotate = turn > 1 && turn <= 8, big = b.Length >= ImageMinBytes;
                    double k = Math.Min(1.0, (double)ImageMaxSide / Math.Max(w, h));
                    if (!rotate && !big) return b;
                    if (!rotate && k >= 1.0 && mime == "image/jpeg") return b;       // a JPEG of a sensible size is not encoded again
                    int nw = Math.Max(1, (int)Math.Round(w * k)), nh = Math.Max(1, (int)Math.Round(h * k));
                    using (var bmp = new System.Drawing.Bitmap(nw, nh, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                    {
                        using (var g = System.Drawing.Graphics.FromImage(bmp))
                        using (var ia = new System.Drawing.Imaging.ImageAttributes())
                        {
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                            g.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                            ia.SetWrapMode(System.Drawing.Drawing2D.WrapMode.TileFlipXY);
                            g.DrawImage(img, new System.Drawing.Rectangle(0, 0, nw, nh), 0, 0, w, h, System.Drawing.GraphicsUnit.Pixel, ia);
                        }
                        if (rotate)
                        {
                            System.Drawing.RotateFlipType[] map = { System.Drawing.RotateFlipType.RotateNoneFlipNone, System.Drawing.RotateFlipType.RotateNoneFlipNone, System.Drawing.RotateFlipType.RotateNoneFlipX, System.Drawing.RotateFlipType.Rotate180FlipNone, System.Drawing.RotateFlipType.Rotate180FlipX, System.Drawing.RotateFlipType.Rotate90FlipX, System.Drawing.RotateFlipType.Rotate90FlipNone, System.Drawing.RotateFlipType.Rotate270FlipX, System.Drawing.RotateFlipType.Rotate270FlipNone };
                            bmp.RotateFlip(map[turn]);
                        }
                        // a photograph: JPEG already, or an opaque picture that PNG could not make small (more than a byte per pixel)
                        bool jpeg = mime == "image/jpeg";
                        byte[] o = null;
                        if (!jpeg && Opaque(bmp))
                        {
                            if (mime == "image/bmp") { o = Encode(bmp, false, 0); jpeg = (double)o.Length / ((double)bmp.Width * bmp.Height) > 1.0; if (jpeg) o = null; }
                            else jpeg = (double)b.Length / ((double)w * h) > 1.0;
                        }
                        if (o == null) o = Encode(bmp, jpeg, 86L);
                        if (!rotate && o.Length > b.Length * 0.85) return b;                // not worth it
                        string was = mime;
                        mime = jpeg ? "image/jpeg" : "image/png";
                        r.Notes.Add("picture " + name + ": " + w + " x " + h + " px, " + Kb(b.Length) + " -> " + bmp.Width + " x " + bmp.Height + " px " + (jpeg ? "JPEG" : "PNG") + ", " + Kb(o.Length) + (rotate ? " (turned upright)" : "") + (was != mime ? " (was " + was.Substring(6).ToUpperInvariant() + ")" : ""));
                        return o;
                    }
                }
            }
            catch { return b; }                                                            // a picture the system cannot decode goes in as it is
        }

        static readonly Regex CssUrlRx = new Regex(@"url\(\s*(?:""([^""]*)""|'([^']*)'|([^)""']*))\s*\)", IC);

        static string CssUrls(string css, string baseDir, PrepResult r)
        {
            return CssUrlRx.Replace(css, delegate(Match m)
            {
                string u = m.Groups[1].Success ? m.Groups[1].Value : (m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Value);
                string d = DataUri(u, baseDir, r, "css url");
                return d == null ? m.Value : "url(\"" + d + "\")";
            });
        }

        static string CssImports(string css, string baseDir, PrepResult r, int depth)
        {
            return Regex.Replace(css, @"@import\s+(?:url\(\s*)?[""']?([^""')\s;]+)[""']?\s*\)?[^;]*;", delegate(Match m)
            {
                string u = m.Groups[1].Value;
                if (IsExternal(u)) { if (!r.External.Contains(u)) r.External.Add(u); return m.Value; }
                string p = LocalPath(u, baseDir);
                if (p == null || !File.Exists(p) || depth > 4) { r.Missing.Add(u + " (@import)"); return m.Value; }
                string sub = ReadText(p);
                string dir = Path.GetDirectoryName(p);
                r.Notes.Add("inlined style sheet " + Path.GetFileName(p));
                return CssUrls(CssImports(sub, dir, r, depth + 1), dir, r);
            }, IC);
        }

        static string MarkupResources(string markup, string baseDir, PrepResult r)
        {
            // style sheets
            markup = Regex.Replace(markup, @"<link\b[^>]*>", delegate(Match m)
            {
                string tag = m.Value;
                string rel = (Attr(tag, "rel") ?? "").ToLowerInvariant();
                string href = Attr(tag, "href");
                if (rel.Contains("icon") || rel == "manifest" || rel == "preload" || rel == "prefetch" || rel == "preconnect" || rel == "dns-prefetch") return "";
                if (!rel.Contains("stylesheet") || href == null) return tag;
                if (IsExternal(href)) { if (!r.External.Contains(href.Trim())) r.External.Add(href.Trim()); return tag; }
                string p = LocalPath(href, baseDir);
                if (p == null || !File.Exists(p)) { r.Missing.Add(href + " (style sheet)"); return tag; }
                string css = ReadText(p);
                string dir = Path.GetDirectoryName(p);
                css = CssUrls(CssImports(css, dir, r, 0), dir, r);
                r.Notes.Add("inlined style sheet " + Path.GetFileName(p));
                string media = Attr(tag, "media");
                return "<style" + (string.IsNullOrEmpty(media) ? "" : " media=\"" + media + "\"") + ">\n" + css.Replace("</style", "<\\/style") + "\n</style>";
            }, IC);
            // src / poster / href of media elements
            markup = Regex.Replace(markup, @"<(img|source|video|audio|input|image|embed|track)\b[^>]*>", delegate(Match m)
            {
                string tag = m.Value;
                foreach (string an in new[] { "src", "poster", "xlink:href", "href" })
                {
                    string v = Attr(tag, an);
                    if (v == null) continue;
                    string d = DataUri(v, baseDir, r, "<" + m.Groups[1].Value.ToLowerInvariant() + " " + an + ">");
                    if (d == null) continue;
                    tag = Regex.Replace(tag, @"(\s" + Regex.Escape(an) + @"\s*=\s*)(?:""[^""]*""|'[^']*'|[^\s>]+)", delegate(Match a) { return a.Groups[1].Value + "\"" + d + "\""; }, IC);
                }
                if (Attr(tag, "srcset") != null) tag = Regex.Replace(tag, @"\ssrcset\s*=\s*(?:""[^""]*""|'[^']*')", "", IC);
                return tag;
            }, IC);
            // style attributes
            markup = Regex.Replace(markup, @"(\sstyle\s*=\s*)(""[^""]*""|'[^']*')", delegate(Match m)
            {
                string v = m.Groups[2].Value;
                if (v.IndexOf("url(", StringComparison.OrdinalIgnoreCase) < 0) return m.Value;
                char q = v[0];
                string inner = CssUrlRx.Replace(v.Substring(1, v.Length - 2), delegate(Match u)
                {
                    string uu = u.Groups[1].Success ? u.Groups[1].Value : (u.Groups[2].Success ? u.Groups[2].Value : u.Groups[3].Value);
                    string d = DataUri(uu, baseDir, r, "style attribute");
                    return d == null ? u.Value : "url(" + d + ")";
                });
                return m.Groups[1].Value + q + inner + q;
            }, IC);
            return markup;
        }

        // ------------------------------------------------------------------ CSS custom properties

        class VarDecl { public string Name, Value, Selector; public int Depth, Start, End; public Seg Owner; }

        static string BlankCssComments(string css)
        {
            return Regex.Replace(css, @"/\*[\s\S]*?\*/", delegate(Match m) { return Regex.Replace(m.Value, @"[^\n]", " "); });
        }

        static void FindVarDecls(Seg seg, List<VarDecl> outList)
        {
            string css = BlankCssComments(seg.Text);
            var stack = new List<string>();
            int segStart = 0; char quote = '\0'; int paren = 0;
            for (int i = 0; i < css.Length; i++)
            {
                char c = css[i];
                if (quote != '\0') { if (c == '\\') i++; else if (c == quote) quote = '\0'; continue; }
                if (c == '"' || c == '\'') { quote = c; continue; }
                if (c == '(') { paren++; continue; }
                if (c == ')') { if (paren > 0) paren--; continue; }
                if (paren > 0) continue;
                if (c == '{') { stack.Add(css.Substring(segStart, i - segStart).Trim()); segStart = i + 1; }
                else if (c == ';' || c == '}')
                {
                    if (stack.Count > 0)
                    {
                        string decl = css.Substring(segStart, i - segStart);
                        Match m = Regex.Match(decl, @"^\s*(--[\w-]+)\s*:([\s\S]*)$");
                        if (m.Success) outList.Add(new VarDecl { Name = m.Groups[1].Value, Value = m.Groups[2].Value.Trim(), Selector = stack[stack.Count - 1], Depth = stack.Count, Start = segStart, End = (c == ';' ? i + 1 : i), Owner = seg });
                    }
                    if (c == '}' && stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                    segStart = i + 1;
                }
            }
        }

        static bool RootSelector(string sel)
        {
            foreach (string part in sel.Split(','))
            {
                string p = part.Trim().ToLowerInvariant();
                if (p != ":root" && p != "html" && p != "body") return false;
            }
            return sel.Trim().Length > 0;
        }

        static readonly Regex VarRx = new Regex(@"var\(\s*(--[\w-]+)\s*", IC);

        // Replaces var(--name[, fallback]) using the map; unknown names keep their fallback, or stay as they are.
        static string ReplaceVars(string text, Dictionary<string, string> map, int depth)
        {
            if (depth > 12) return text;
            var sb = new StringBuilder();
            int pos = 0;
            while (true)
            {
                Match m = VarRx.Match(text, pos);
                if (!m.Success) { sb.Append(text, pos, text.Length - pos); break; }
                int start = m.Index;
                int i = start + m.Length;
                int level = 1; int argStart = -1;
                if (i < text.Length && text[i] == ',') { argStart = i + 1; }
                int j = i;
                for (; j < text.Length; j++)
                {
                    if (text[j] == '(') level++;
                    else if (text[j] == ')') { level--; if (level == 0) break; }
                }
                if (j >= text.Length) { sb.Append(text, pos, text.Length - pos); break; }
                string name = m.Groups[1].Value;
                string fallback = argStart >= 0 ? text.Substring(argStart, j - argStart).Trim() : null;
                sb.Append(text, pos, start - pos);
                if (map.ContainsKey(name)) sb.Append(ReplaceVars(map[name], map, depth + 1));
                else if (fallback != null) sb.Append(ReplaceVars(fallback, map, depth + 1));
                else sb.Append(text, start, j + 1 - start);
                pos = j + 1;
            }
            return sb.ToString();
        }

        static void InlineCssVars(List<Seg> segs, PrepResult r)
        {
            var decls = new List<VarDecl>();
            bool anyUse = false, jsTouches = false;
            foreach (var s in segs)
            {
                if (s.Kind == 3) { FindVarDecls(s, decls); if (Regex.IsMatch(s.Text, @"var\(\s*--", IC)) anyUse = true; }
                else if (s.Kind == 0) { if (Regex.IsMatch(s.Text, @"var\(\s*--", IC)) anyUse = true; }
                else if (s.Kind == 2) { if (Regex.IsMatch(s.Text, @"(setProperty|getPropertyValue)\(\s*[""']--")) jsTouches = true; }
            }
            if (decls.Count == 0 && !anyUse) return;
            if (jsTouches) { r.Notes.Add("CSS variables are read or set by script: not resolved automatically"); return; }
            var count = new Dictionary<string, int>();
            foreach (var d in decls) { int n; count.TryGetValue(d.Name, out n); count[d.Name] = n + 1; }
            var map = new Dictionary<string, string>();
            var good = new List<VarDecl>();
            var kept = new List<string>();
            foreach (var d in decls)
            {
                if (count[d.Name] == 1 && d.Depth == 1 && RootSelector(d.Selector)) { map[d.Name] = d.Value; good.Add(d); }
                else if (!kept.Contains(d.Name)) kept.Add(d.Name);
            }
            // remove the resolved declarations (back to front inside each style block)
            good.Sort(delegate(VarDecl a, VarDecl b) { return b.Start.CompareTo(a.Start); });
            foreach (var d in good) d.Owner.Text = d.Owner.Text.Remove(d.Start, d.End - d.Start);
            foreach (var s in segs)
            {
                if (s.Kind == 3) s.Text = ReplaceVars(s.Text, map, 0);
                else if (s.Kind == 0 && s.Text.IndexOf("var(", StringComparison.OrdinalIgnoreCase) >= 0)
                    s.Text = Regex.Replace(s.Text, @"(\sstyle\s*=\s*)(""[^""]*""|'[^']*')", delegate(Match m) { return m.Groups[1].Value + ReplaceVars(m.Groups[2].Value, map, 0); }, IC);
            }
            if (good.Count > 0) r.Notes.Add("CSS variables written out as plain values: " + good.Count + (kept.Count > 0 ? " (the others could not be, see next note)" : ""));
            if (kept.Count > 0) r.Notes.Add("CSS variables NOT written out, because they are defined more than once (e.g. a theme override) or outside :root/html/body. Replace them by hand: " + string.Join(", ", kept.ToArray()));
        }

        // ------------------------------------------------------------------ prepare

        // html: page source. baseDir: folder for relative references. harnessHead / harnessTail: one-line blocks
        // that go right after <head> and right before </body>. polyfills: script text added when the page uses
        // functions IE11 lacks ("" = never add).
        public static PrepResult Prepare(string html, string baseDir, string harnessHead, string harnessTail, string polyfills)
        {
            return Finish(Begin(html, baseDir), harnessHead, harnessTail, polyfills, "", "");
        }

        static readonly Regex StyleAttrRx = new Regex(@"(\sstyle\s*=\s*)(""[^""]*""|'[^']*')", IC);

        static bool IsJsType(string open)
        {
            string type = (Attr(open, "type") ?? "").Trim().ToLowerInvariant();
            return type == "" || type.Contains("javascript") || type.Contains("ecmascript") || type == "module";
        }

        static string AttrDecode(string s)
        {
            if (s.IndexOf('&') < 0) return s;
            return s.Replace("&quot;", "\"").Replace("&#34;", "\"").Replace("&#39;", "'").Replace("&apos;", "'").Replace("&lt;", "<").Replace("&gt;", ">").Replace("&amp;", "&");
        }

        static string AttrEncode(string s) { return s.Replace("&", "&amp;").Replace("\"", "&quot;"); }

        // Step 1: pulls the local files into the page and lists what a converter may want to rewrite.
        public static PrepState Begin(string html, string baseDir)
        {
            var st = new PrepState();
            var r = st.R;
            html = html.Replace("\r\n", "\n").Replace('\r', '\n');
            html = Regex.Replace(html, @"[\x00-\x08\x0B\x0C\x0E-\x1F]", "");
            html = StripHarness(html);
            var segs = Split(html);
            var outSegs = new List<Seg>();
            foreach (var s in segs)
            {
                if (s.Kind == 1)
                {
                    if (Regex.IsMatch(s.Text, @"^<!--\s*saved from url=", IC)) continue;
                    outSegs.Add(s);
                }
                else if (s.Kind == 2)
                {
                    string src = Attr(s.Open, "src");
                    if (src != null && !IsExternal(src))
                    {
                        string p = LocalPath(src, baseDir);
                        if (p != null && File.Exists(p))
                        {
                            string code = ReadText(p);
                            r.InlinedBytes += code.Length;
                            s.Text = "\n" + Regex.Replace(code, @"</script", "<\\/script", IC) + "\n";
                            s.Open = Regex.Replace(s.Open, @"\ssrc\s*=\s*(?:""[^""]*""|'[^']*'|[^\s>]+)", "", IC);
                            r.Notes.Add("inlined script " + Path.GetFileName(p));
                        }
                        else r.Missing.Add(src + " (script)");
                    }
                    else if (src != null) { if (!r.External.Contains(src.Trim())) r.External.Add(src.Trim()); }
                    outSegs.Add(s);
                }
                else if (s.Kind == 3)
                {
                    s.Text = CssUrls(CssImports(s.Text, baseDir, r, 0), baseDir, r);
                    outSegs.Add(s);
                }
                else
                {
                    string t = s.Text;
                    t = Regex.Replace(t, @"<!DOCTYPE[^>]*>", "", IC);
                    t = Regex.Replace(t, @"<meta\b[^>]*\bcharset\s*=[^>]*>", "", IC);
                    t = Regex.Replace(t, @"<meta\b[^>]*http-equiv\s*=\s*[""']?(content-type|x-ua-compatible)[^>]*>", "", IC);
                    t = MarkupResources(t, baseDir, r);
                    s.Text = t;
                    outSegs.Add(s);
                }
            }
            // style blocks produced from <link> live inside markup text now: split again so they are handled as CSS
            segs = Split(Join(outSegs));
            st.Segs = segs;
            st.Doc = Join(segs).Trim();
            var scripts = new List<string>(); var mods = new List<bool>(); var styles = new List<string>(); var inl = new List<string>();
            foreach (var s in segs)
            {
                if (s.Kind == 2 && Attr(s.Open, "src") == null && IsJsType(s.Open) && s.Text.Trim().Length > 0)
                {
                    st.ScriptSegs.Add(s); scripts.Add(s.Text);
                    mods.Add((Attr(s.Open, "type") ?? "").Trim().ToLowerInvariant() == "module");
                }
                else if (s.Kind == 3)
                {
                    st.StyleSegs.Add(s);
                    string media = (Attr(s.Open, "media") ?? "").Trim();
                    string ml = media.ToLowerInvariant();
                    styles.Add(ml == "" || ml == "all" || ml == "screen" ? s.Text : "@media " + media + " {\n" + s.Text + "\n}");
                }
                else if (s.Kind == 0)
                {
                    foreach (Match m in StyleAttrRx.Matches(s.Text)) { string v = m.Groups[2].Value; inl.Add(AttrDecode(v.Substring(1, v.Length - 2))); }
                }
            }
            st.Scripts = scripts.ToArray(); st.ScriptIsModule = mods.ToArray(); st.Styles = styles.ToArray(); st.InlineStyles = inl.ToArray();
            // a first look with the rules, so that the caller can tell whether a conversion is worth starting
            // (on a copy in which the variables the simple resolver handles have been written out: they need no conversion)
            var probe = new List<Issue>();
            var copy = Split(st.Doc);
            InlineCssVars(copy, new PrepResult());
            Lint("<!DOCTYPE html>\n" + Join(copy), probe, null);
            foreach (var i in probe)
            {
                if (i.Severity != "error") continue;
                if (i.Rule.StartsWith("js-") && i.Rule != "js-module") st.NeedsJs = true;
                else if (i.Rule.StartsWith("css-")) st.NeedsCss = true;
                else if (i.Rule == "html-module" || i.Rule == "js-module") st.NeedsJs = true;
            }
            st.OptOut = Regex.IsMatch(html, @"<!--\s*ppt-html-convert\s*:\s*(off|no|false)\s*-->", IC);
            return st;
        }

        // Step 2 (optional): takes what the converter produced. A null entry means "leave as it is".
        // inlineV: for style attributes, the JSON the layout adapter reads (data-ih-v). tableJson: its rule table.
        public static void ApplyConversion(PrepState st, string[] scripts, string css, string[] inlineStyles, string[] inlineV, string tableJson, string[] allow, string[] notes)
        {
            var r = st.R;
            if (scripts != null)
            {
                for (int i = 0; i < st.ScriptSegs.Count && i < scripts.Length; i++)
                {
                    if (string.IsNullOrEmpty(scripts[i])) continue;
                    st.ScriptSegs[i].Text = "\n" + Regex.Replace(scripts[i], @"</script", "<\\/script", IC) + "\n";
                    st.ScriptsConverted = true;
                }
            }
            // module scripts without imports: IE11 ignores type="module", so they become plain scripts that run
            // where module scripts run (after the document has been parsed), in a scope of their own
            for (int i = 0; i < st.ScriptSegs.Count; i++)
            {
                if (!st.ScriptIsModule[i]) continue;
                var s = st.ScriptSegs[i];
                if (Regex.IsMatch(CodeOnlyJs(s.Text), @"\bimport\s*[\w{*(""']|\bexport\s+(default|function|class|\{|var|const|let)")) continue;
                string open = Regex.Replace(s.Open, @"\stype\s*=\s*(?:""[^""]*""|'[^']*'|[^\s>]+)", "", IC);
                st.Deferred += open + "(function () { \"use strict\";" + s.Text + "}());" + s.Close + "\n";
                s.Kind = 1; s.Text = "";
                st.ScriptsConverted = true;
                r.Notes.Add("converted: a module script without imports now runs as a plain script at the end of the page");
            }
            if (!string.IsNullOrEmpty(css) && st.StyleSegs.Count > 0)
            {
                st.StyleSegs[0].Open = "<style>";
                st.StyleSegs[0].Text = "\n" + css.Replace("</style", "<\\/style") + "\n";
                for (int i = 1; i < st.StyleSegs.Count; i++) { st.StyleSegs[i].Kind = 1; st.StyleSegs[i].Text = ""; }
            }
            if (inlineStyles != null)
            {
                int n = 0;
                foreach (var s in st.Segs)
                {
                    if (s.Kind != 0) continue;
                    s.Text = StyleAttrRx.Replace(s.Text, delegate(Match m)
                    {
                        int k = n++;
                        if (k >= inlineStyles.Length || string.IsNullOrEmpty(inlineStyles[k])) return m.Value;
                        string extra = (inlineV != null && k < inlineV.Length && !string.IsNullOrEmpty(inlineV[k])) ? " data-ih-v='" + inlineV[k].Replace("&", "&amp;").Replace("'", "&#39;") + "'" : "";
                        return m.Groups[1].Value + "\"" + AttrEncode(inlineStyles[k]) + "\"" + extra;
                    });
                }
            }
            st.TableJson = tableJson ?? "";
            st.Allow = allow ?? new string[0];
            st.Converted = true;
            if (notes != null) foreach (var x in notes) r.Notes.Add("converted: " + x);
        }

        // Adds CSS that is to sit in a style block of its own at the start of the head (used for embedded fonts).
        public static void AddCss(PrepState st, string css) { st.ExtraCss += css; }

        static string OneLine(string script) { return Regex.Replace(script.Replace("\r", ""), @"\n\s*", " ").Replace("</script", "<\\/script"); }

        // Step 3: the document skeleton, the harness blocks, the check against the IE11 rules, the helper scripts.
        // heavyPolyfills: the ES2015+ library (core-js), added when scripts were rewritten. adapter: the page-side
        // layout script, added when the converter left a table for it.
        public static PrepResult Finish(PrepState st, string harnessHead, string harnessTail, string polyfills, string heavyPolyfills, string adapter)
        {
            var r = st.R;
            var segs = st.Segs;
            if (!st.Converted) InlineCssVars(segs, r);
            string doc = Join(segs).Trim();

            // document skeleton
            segs = Split(doc);
            bool hasHtml = false, hasHead = false, hasBody = false;
            foreach (var s in segs)
            {
                if (s.Kind != 0) continue;
                if (Regex.IsMatch(s.Text, @"<html\b", IC)) hasHtml = true;
                if (Regex.IsMatch(s.Text, @"<head\b", IC)) hasHead = true;
                if (Regex.IsMatch(s.Text, @"<body\b", IC)) hasBody = true;
            }
            if (!hasHtml && !hasBody && !hasHead) doc = "<html>\n<head>\n</head>\n<body>\n" + doc + "\n</body>\n</html>";
            else
            {
                if (!hasHtml) doc = "<html>\n" + doc + "\n</html>";
                if (!hasHead) doc = ReplaceFirstInMarkup(doc, @"<html\b[^>]*>", delegate(string m) { return m + "\n<head>\n</head>"; });
                if (!hasBody)
                {
                    doc = ReplaceFirstInMarkup(doc, @"</head\s*>", delegate(string m) { return m + "\n<body>"; });
                    doc = ReplaceLastInMarkup(doc, @"</html\s*>", delegate(string m) { return "</body>\n" + m; });
                }
            }
            if (st.Deferred.Length > 0)
            {
                bool put = false;
                doc = ReplaceLastInMarkup(doc, @"</body\s*>", delegate(string m) { put = true; return st.Deferred + m; });
                if (!put) doc += st.Deferred;
            }
            if (st.ExtraCss.Length > 0)
                doc = ReplaceFirstInMarkup(doc, @"<head\b[^>]*>", delegate(string m) { return m + "\n<style>\n" + st.ExtraCss + "</style>"; });
            string head = harnessHead, tail = harnessTail;
            bool wantsAdapter = st.Converted && !string.IsNullOrEmpty(adapter) && ((st.TableJson.Length > 2) || doc.IndexOf("data-ih-v=", StringComparison.Ordinal) >= 0);
            if (wantsAdapter)
                tail = "<script>window.__ihAdapt=" + (st.TableJson.Length > 2 ? st.TableJson.Replace("</", "<\\/") : "[]") + ";</script><script>" + OneLine(adapter) + "</script>" + tail;
            var allow = new List<string>(st.Allow);
            bool heavy = st.ScriptsConverted && !string.IsNullOrEmpty(heavyPolyfills);
            if (heavy) { allow.Add("api-symbol"); allow.Add("api-map-set"); allow.Add("api-url-weakset"); }
            r.Html = Assemble(doc, head, tail);
            r.UsesViewportUnits = Regex.IsMatch(r.Html, @"\d(vw|vh|vmin)\b", IC);
            Lint(r.Html, r.Issues, allow);

            // Functions that IE11 lacks but that can be supplied (find, includes, closest, Promise ...): add them
            // to the page instead of asking for a rewrite. Syntax and CSS cannot be supplied this way.
            var blocks = new StringBuilder();
            if (heavy)
            {
                blocks.Append("<script>" + heavyPolyfills.Replace("</script", "<\\/script") + "</script>");
                r.Notes.Add("converted: added the ES2015+ library functions (core-js) that rewritten scripts rely on");
            }
            if (!string.IsNullOrEmpty(polyfills))
            {
                var used = new List<string>();
                foreach (var i in r.Issues) if (Array.IndexOf(Polyfilled, i.Rule) >= 0 && !used.Contains(i.Rule)) used.Add(i.Rule);
                if (used.Count > 0)
                {
                    blocks.Append("<script>" + OneLine(polyfills) + "</script>");
                    r.Issues.RemoveAll(delegate(Issue i) { return Array.IndexOf(Polyfilled, i.Rule) >= 0; });
                    r.Notes.Add("added the IE11 helper functions (polyfills) because the page uses: " + string.Join(", ", used.ToArray()));
                }
            }
            if (blocks.Length > 0) r.Html = Assemble(doc, head + blocks.ToString(), tail);
            return r;
        }

        // The page for the modern engine (Edge WebView2 in the slide show, any current browser for the web version):
        // the page as it was written, with its local files pulled in and the embedded fonts, plus the harness of
        // that engine. Nothing is converted and nothing is checked against the IE11 rules.
        public static string FinishModern(PrepState st, string harnessHead, string harnessTail)
        {
            string doc = st.Doc.Trim();
            var segs = Split(doc);
            bool hasHtml = false, hasHead = false, hasBody = false;
            foreach (var s in segs)
            {
                if (s.Kind != 0) continue;
                if (Regex.IsMatch(s.Text, @"<html\b", IC)) hasHtml = true;
                if (Regex.IsMatch(s.Text, @"<head\b", IC)) hasHead = true;
                if (Regex.IsMatch(s.Text, @"<body\b", IC)) hasBody = true;
            }
            if (!hasHtml && !hasBody && !hasHead) doc = "<html>\n<head>\n</head>\n<body>\n" + doc + "\n</body>\n</html>";
            else
            {
                if (!hasHtml) doc = "<html>\n" + doc + "\n</html>";
                if (!hasHead) doc = ReplaceFirstInMarkup(doc, @"<html\b[^>]*>", delegate(string m) { return m + "\n<head>\n</head>"; });
                if (!hasBody)
                {
                    doc = ReplaceFirstInMarkup(doc, @"</head\s*>", delegate(string m) { return m + "\n<body>"; });
                    doc = ReplaceLastInMarkup(doc, @"</html\s*>", delegate(string m) { return "</body>\n" + m; });
                }
            }
            if (st.ExtraCss.Length > 0)
                doc = ReplaceFirstInMarkup(doc, @"<head\b[^>]*>", delegate(string m) { return m + "\n<style>\n" + st.ExtraCss + "</style>"; });
            return Assemble(doc, harnessHead, harnessTail);
        }

        // Rules whose findings are covered by assets\polyfills-ie11.txt.
        static readonly string[] Polyfilled = { "api-static", "api-string-array", "api-find", "api-dom", "api-remove", "api-nodelist-foreach", "api-event-ctor", "api-classlist", "api-promise" };

        static string Assemble(string doc, string harnessHead, string harnessTail)
        {
            doc = ReplaceFirstInMarkup(doc, @"<head\b[^>]*>", delegate(string m) { return m + XUa + HeadOpen + harnessHead + HeadClose; });
            bool tailDone = false;
            doc = ReplaceLastInMarkup(doc, @"</body\s*>", delegate(string m) { tailDone = true; return TailOpen + harnessTail + TailClose + m; });
            if (!tailDone) doc = doc + TailOpen + harnessTail + TailClose;
            return "<!DOCTYPE html>\n" + doc + "\n";
        }

        delegate string Rep(string match);

        static string ReplaceFirstInMarkup(string doc, string pattern, Rep rep)
        {
            var segs = Split(doc);
            foreach (var s in segs)
            {
                if (s.Kind != 0) continue;
                Match m = Regex.Match(s.Text, pattern, IC);
                if (!m.Success) continue;
                s.Text = s.Text.Substring(0, m.Index) + rep(m.Value) + s.Text.Substring(m.Index + m.Length);
                break;
            }
            return Join(segs);
        }

        static string ReplaceLastInMarkup(string doc, string pattern, Rep rep)
        {
            var segs = Split(doc);
            for (int i = segs.Count - 1; i >= 0; i--)
            {
                var s = segs[i];
                if (s.Kind != 0) continue;
                MatchCollection mc = Regex.Matches(s.Text, pattern, IC);
                if (mc.Count == 0) continue;
                Match m = mc[mc.Count - 1];
                s.Text = s.Text.Substring(0, m.Index) + rep(m.Value) + s.Text.Substring(m.Index + m.Length);
                break;
            }
            return Join(segs);
        }

        // Removes the harness blocks (and the X-UA meta that goes with them) from a prepared document.
        public static string StripHarness(string doc)
        {
            doc = Regex.Replace(doc, Regex.Escape(XUa) + Regex.Escape(HeadOpen) + @"[\s\S]*?" + Regex.Escape(HeadClose), "");
            doc = Regex.Replace(doc, Regex.Escape(HeadOpen) + @"[\s\S]*?" + Regex.Escape(HeadClose), "");
            doc = Regex.Replace(doc, Regex.Escape(TailOpen) + @"[\s\S]*?" + Regex.Escape(TailClose), "");
            return doc;
        }

        // Swaps the harness of a prepared document for another one (e.g. the web viewer's).
        public static string SwapHarness(string doc, string harnessHead, string harnessTail)
        {
            doc = Regex.Replace(doc, Regex.Escape(HeadOpen) + @"[\s\S]*?" + Regex.Escape(HeadClose), delegate(Match m) { return HeadOpen + harnessHead + HeadClose; });
            doc = Regex.Replace(doc, Regex.Escape(TailOpen) + @"[\s\S]*?" + Regex.Escape(TailClose), delegate(Match m) { return TailOpen + harnessTail + TailClose; });
            return doc;
        }

        // ------------------------------------------------------------------ lint

        static int LineOf(string text, int index)
        {
            int n = 1;
            for (int i = 0; i < index && i < text.Length; i++) if (text[i] == '\n') n++;
            return n;
        }

        static string Snip(string text, int index)
        {
            int a = text.LastIndexOf('\n', Math.Max(0, Math.Min(index, text.Length - 1)));
            int b = text.IndexOf('\n', Math.Min(index, text.Length));
            if (b < 0) b = text.Length;
            string line = text.Substring(a + 1, b - a - 1).Trim();
            if (line.Length > 110)
            {
                int rel = index - (a + 1);
                int from = Math.Max(0, Math.Min(line.Length - 110, rel - 40));
                line = (from > 0 ? "..." : "") + line.Substring(from, Math.Min(110, line.Length - from)) + "...";
            }
            return line;
        }

        // Blanks strings, comments and regular expression literals of JavaScript, keeping length and line breaks.
        // A template literal is kept as a pair of back ticks so that the rule for it can fire.
        static string CodeOnlyJs(string js)
        {
            var sb = new StringBuilder(js);
            int i = 0; int n = js.Length; char prevSig = '\0'; string prevWord = "";
            while (i < n)
            {
                char c = js[i];
                if (c == '/' && i + 1 < n && js[i + 1] == '/')
                {
                    while (i < n && js[i] != '\n') { sb[i] = ' '; i++; }
                    continue;
                }
                if (c == '/' && i + 1 < n && js[i + 1] == '*')
                {
                    int e = js.IndexOf("*/", i + 2, StringComparison.Ordinal);
                    if (e < 0) e = n - 2;
                    for (int k = i; k < e + 2 && k < n; k++) if (js[k] != '\n') sb[k] = ' ';
                    i = e + 2; continue;
                }
                if (c == '"' || c == '\'' || c == '`')
                {
                    int k = i + 1;
                    while (k < n && js[k] != c)
                    {
                        if (js[k] == '\\') { sb[k] = ' '; k++; if (k < n && js[k] != '\n') sb[k] = ' '; k++; continue; }
                        if (js[k] == '\n' && c != '`') break;
                        if (js[k] != '\n') sb[k] = ' ';
                        k++;
                    }
                    i = k + 1; prevSig = c; prevWord = ""; continue;
                }
                if (c == '/')
                {
                    bool regex = prevSig == '\0' || "(,=:[!&|?{};+-*%<>~^".IndexOf(prevSig) >= 0 || prevWord == "return" || prevWord == "typeof" || prevWord == "case" || prevWord == "in" || prevWord == "of" || prevWord == "void" || prevWord == "delete";
                    if (regex)
                    {
                        int k = i + 1; bool cls = false;
                        while (k < n && js[k] != '\n')
                        {
                            if (js[k] == '\\') { sb[k] = ' '; k++; if (k < n) sb[k] = ' '; k++; continue; }
                            if (js[k] == '[') cls = true;
                            else if (js[k] == ']') cls = false;
                            else if (js[k] == '/' && !cls) break;
                            sb[k] = ' ';
                            k++;
                        }
                        i = k + 1; prevSig = 'x'; prevWord = ""; continue;
                    }
                }
                if (char.IsLetterOrDigit(c) || c == '_' || c == '$')
                {
                    int k = i;
                    while (k < n && (char.IsLetterOrDigit(js[k]) || js[k] == '_' || js[k] == '$')) k++;
                    prevWord = js.Substring(i, k - i); prevSig = 'x'; i = k; continue;
                }
                if (!char.IsWhiteSpace(c)) { prevSig = c; prevWord = ""; }
                i++;
            }
            return sb.ToString();
        }

        class Rule { public string Sev, Id, Pattern, Msg; public Rule(string s, string id, string p, string m) { Sev = s; Id = id; Pattern = p; Msg = m; } }

        static readonly Rule[] JsRules = new[]
        {
            new Rule("error", "js-arrow", @"=>", "Arrow functions are a syntax error in IE11: write function(){...}."),
            new Rule("error", "js-template", @"`", "Template literals (back ticks) are a syntax error in IE11: join strings with +."),
            new Rule("error", "js-let-const", @"\b(let|const)\s+[\w$\[{]", "let / const: use var (inside loops, pass the value through a function to keep it)."),
            new Rule("error", "js-class", @"\bclass\s+[A-Za-z_$][\w$]*\s*(extends\s+[\w$.]+\s*)?\{", "class syntax is not supported: use a constructor function and prototype."),
            new Rule("error", "js-async", @"\basync\s+function|\basync\s*\(|\bawait\s", "async / await are not supported."),
            new Rule("error", "js-spread", @"\.\.\.", "Spread / rest (...) is not supported: use apply, concat or arguments."),
            new Rule("error", "js-for-of", @"\bfor\s*\([^)]*\sof\s", "for...of is not supported: use an index loop."),
            new Rule("error", "js-optional-chain", @"\?\.(?=[A-Za-z_$(\[])", "Optional chaining (?.) is not supported."),
            new Rule("error", "js-nullish", @"\?\?", "The ?? operator is not supported: use an explicit null check."),
            new Rule("error", "js-exponent", @"\*\*", "The ** operator is not supported: use Math.pow."),
            new Rule("error", "js-destructure", @"\bvar\s*[\[{]", "Destructuring is not supported."),
            new Rule("error", "js-module", @"\bimport\s+[\w{*""']|\bexport\s+(default|function|class|\{|var)", "ES modules are not supported: use one plain script."),
            new Rule("error", "api-fetch", @"\bfetch\s*\(", "fetch does not exist in IE11 (XMLHttpRequest does; pages run from a local file, so keep data inline)."),
            new Rule("error", "api-promise", @"\bPromise\b", "Promise does not exist in IE11: use callbacks or include a polyfill."),
            new Rule("error", "api-symbol", @"\bSymbol\b", "Symbol does not exist in IE11."),
            new Rule("error", "api-static", @"\b(Array\.from|Array\.of|Object\.assign|Object\.entries|Object\.values|Object\.fromEntries|Number\.isFinite|Number\.isNaN|Number\.isInteger|Number\.parseFloat|Number\.parseInt|Math\.sign|Math\.trunc|Math\.log2|Math\.log10|Math\.hypot|Math\.cbrt)\b", "This built-in function does not exist in IE11: write it out or add a small polyfill."),
            new Rule("error", "api-string-array", @"\.(includes|startsWith|endsWith|padStart|padEnd|repeat|findIndex|flatMap|flat|at)\s*\(", "This string / array method does not exist in IE11 (use indexOf, a loop, etc.)."),
            new Rule("error", "api-find", @"\.find\s*\(", "Array.prototype.find does not exist in IE11: use a loop or filter(...)[0]."),
            new Rule("error", "api-dom", @"\.(closest|append|prepend|before|after|replaceWith|replaceChildren|toggleAttribute|getAttributeNames)\s*\(", "This DOM method does not exist in IE11 (use appendChild / insertBefore / parentNode loops)."),
            new Rule("error", "api-remove", @"\.remove\s*\(\s*\)", "element.remove() does not exist in IE11: use el.parentNode.removeChild(el)."),
            new Rule("error", "api-nodelist-foreach", @"(querySelectorAll\s*\([^)]*\)|\.childNodes|\.children)\s*\.forEach\s*\(", "NodeList / HTMLCollection has no forEach in IE11: use an index loop."),
            new Rule("error", "api-event-ctor", @"\bnew\s+(CustomEvent|Event)\s*\(", "Event constructors do not work in IE11: use document.createEvent."),
            new Rule("error", "api-event-ctor-input", @"\bnew\s+(MouseEvent|KeyboardEvent|PointerEvent|InputEvent|TouchEvent)\s*\(", "This event constructor does not work in IE11: use document.createEvent and init...Event."),
            new Rule("error", "api-missing-class", @"\bnew\s+(IntersectionObserver|ResizeObserver|Proxy|AbortController|TextEncoder|TextDecoder)\b", "This class does not exist in IE11."),
            new Rule("error", "api-url-weakset", @"\bnew\s+(URLSearchParams|URL|WeakSet)\b", "This class does not exist in IE11."),
            new Rule("warn", "css-var-js", @"(setProperty|getPropertyValue)\(\s*[""']--", "CSS variables are settled when the page is built: reading or changing them from script has no effect in IE11. Toggle classes instead."),
            new Rule("warn", "api-storage", @"\b(localStorage|sessionStorage)\b", "Web storage is not available for pages opened from a local file in IE11: keep state in variables."),
            new Rule("warn", "api-classlist", @"\.classList\.(add|remove)\s*\([^)]*,|\.classList\.toggle\s*\([^)]*,|\.classList\.replace\s*\(", "IE11 classList: one class per add/remove call, toggle has no second argument, no replace."),
            new Rule("warn", "api-key", @"\.key\s*===?\s*[""']|\.code\s*===?\s*[""']|[""']\s*===?\s*[\w$.]+\.(key|code)\b", "event.key uses old names in IE11 (Left, Right, Up, Down, Esc, Spacebar) and event.code does not exist: compare event.keyCode."),
            new Rule("warn", "api-touch", @"\b(touchstart|touchmove|touchend)\b", "Touch events do not exist in IE11: also handle mouse (or pointer) events."),
            new Rule("warn", "api-map-set", @"\bnew\s+(Map|Set|WeakMap)\s*\(\s*[^)\s]", "IE11 Map / Set ignore constructor arguments and have no iterators: fill them with set() / add()."),
        };

        static readonly Rule[] CssRules = new[]
        {
            new Rule("error", "css-var", @"var\(\s*--", "CSS variables are not supported in IE11: write the value itself. (They are resolved automatically only when each one is defined once in :root, html or body.)"),
            new Rule("error", "css-var-decl", @"(?<![\w-])--[\w-]+\s*:", "CSS variable definition: not supported in IE11."),
            new Rule("error", "css-grid", @"display\s*:\s*(inline-)?grid|(?<![\w-])grid-(template|area|column|row|auto|gap)", "CSS grid is not supported in IE11: use floats, inline-block, tables, absolute positions or flexbox."),
            new Rule("error", "css-gap", @"(?<![\w-])(gap|row-gap|column-gap)\s*:", "gap is not supported in IE11 (also not for flexbox): use margins. (column-gap of multi-column layout is fine.)"),
            new Rule("warn", "css-flex", @"display\s*:\s*(inline-)?flex", "Flexbox works in IE11 with quirks: give flex items an explicit flex: 1 1 auto or a width, avoid min-height on the container, no gap. Check the preview picture."),
            new Rule("error", "css-sticky", @"position\s*:\s*sticky", "position: sticky is not supported in IE11."),
            new Rule("error", "css-object-fit", @"(?<![\w-])object-(fit|position)\s*:", "object-fit is not supported in IE11: use a background-image with background-size."),
            new Rule("error", "css-aspect-ratio", @"(?<![\w-])aspect-ratio\s*:", "aspect-ratio is not supported in IE11: set width and height."),
            new Rule("error", "css-math-fn", @"(?<![\w-])(clamp|min|max)\s*\(", "clamp() / min() / max() are not supported in IE11: use a fixed value or calc()."),
            new Rule("error", "css-inset", @"(?<![\w-])inset\s*:", "inset is not supported in IE11: write top / right / bottom / left."),
            new Rule("error", "css-selector", @":(is|where|has|focus-visible|focus-within)\b", "This selector is not supported in IE11 and makes the whole rule be ignored."),
            new Rule("error", "css-backdrop", @"backdrop-filter\s*:", "backdrop-filter is not supported in IE11."),
            new Rule("warn", "css-conic", @"conic-gradient\(", "conic-gradient is not supported in IE11 (the background is not drawn): use an SVG or a linear / radial gradient."),
            new Rule("error", "css-filter", @"(?<![\w-])filter\s*:\s*(blur|brightness|contrast|drop-shadow|grayscale|hue-rotate|invert|opacity|saturate|sepia)\s*\(", "CSS filter functions are not supported in IE11."),
            new Rule("error", "css-blend", @"(mix-blend-mode|background-blend-mode)\s*:", "Blend modes are not supported in IE11."),
            new Rule("error", "css-clip-path", @"(?<![\w-])clip-path\s*:", "clip-path is not supported in IE11 on HTML elements."),
            new Rule("warn", "css-supports", @"@supports\b", "@supports is not understood by IE11: everything inside the block is ignored."),
            new Rule("error", "css-at-rule", @"@(container|layer|property)\b", "This at-rule is not supported in IE11."),
            new Rule("warn", "css-unit", @"\d(vmax|svh|svw|dvh|dvw|lvh|lvw)\b", "This unit is missing or unreliable in IE11 (vmax, dynamic viewport units): use vw / vh / px / rem."),
            new Rule("warn", "css-hex-alpha", @"#[0-9a-fA-F]{8}(?![0-9a-fA-F])", "8-digit hex colours are not supported in IE11: use rgba()."),
            new Rule("warn", "css-user-select", @"(?<![\w-])user-select\s*:", "user-select needs the -ms- prefix in IE11 (add -ms-user-select)."),
            new Rule("warn", "perf-transition-all", @"transition(-property)?\s*:\s*all\b", "transition: all also animates sizes and shadows, which the slide show engine repaints slowly: name the properties (opacity, transform, background-color)."),
            new Rule("warn", "perf-shadow-blur", @"box-shadow\s*:[^;{}]*?(?<![\w.-])-?\d+(?:\.\d+)?(?:px)?\s+-?\d+(?:\.\d+)?(?:px)?\s+(?:[4-9]\d|\d{3,})(?:\.\d+)?px", "A shadow blurred by 40px or more is repainted whenever something near it changes: keep the blur under 30px (or drop the shadow) on anything that moves, fades or sits next to moving content."),
            new Rule("warn", "css-place", @"(?<![\w-])place-(items|content|self)\s*:", "place-* shorthands are not supported in IE11."),
        };

        static readonly Rule[] HtmlRules = new[]
        {
            new Rule("error", "html-module", @"<script\b[^>]*type\s*=\s*[""']?module", "Module scripts do not run in IE11."),
            new Rule("error", "html-template", @"<template\b", "<template> is not supported in IE11 (its content is shown as normal markup)."),
            new Rule("warn", "html-element", @"<(details|dialog|picture|slot|datalist|meter|output)\b", "This element has no built-in behaviour in IE11."),
            new Rule("warn", "html-input-type", @"<input\b[^>]*type\s*=\s*[""']?(date|color|datetime-local|month|week|time)\b", "This input type is a plain text box in IE11."),
            new Rule("warn", "html-range", @"<input\b[^>]*type\s*=\s*[""']?range\b", "IE11 fires 'change' (not 'input') while a range slider is dragged: listen to both events."),
            new Rule("warn", "html-media", @"<(video|audio)\b", "Video / audio depends on the codecs of the PC and is missing from the static picture; test on the presentation PC."),
        };

        static void Apply(Rule[] rules, string scan, string original, int offsetLine, List<Issue> issues, HashSet<string> allow)
        {
            foreach (var rule in rules)
            {
                if (allow.Contains(rule.Id)) continue;
                if (rule.Id == "css-user-select" && scan.IndexOf("-ms-user-select", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                MatchCollection mc = Regex.Matches(scan, rule.Pattern);
                if (mc.Count == 0) continue;
                var lines = new List<int>();
                int count = 0; string snippet = null;
                foreach (Match m in mc)
                {
                    if (rule.Id == "css-gap" && Regex.IsMatch(Snip(original, m.Index), @"column-count|columns\s*:", IC)) continue;
                    count++;
                    int ln = offsetLine + LineOf(scan, m.Index) - 1;
                    if (!lines.Contains(ln)) lines.Add(ln);
                    if (snippet == null) snippet = Snip(original, m.Index);
                }
                if (count == 0) continue;
                // one entry per rule and place in the page (style block, script block ...): all line numbers, first line shown
                var sb = new StringBuilder();
                for (int k = 0; k < lines.Count && k < 12; k++) { if (k > 0) sb.Append(", "); sb.Append(lines[k]); }
                if (lines.Count > 12) sb.Append(", ... (" + lines.Count + " lines)");
                issues.Add(new Issue { Severity = rule.Sev, Rule = rule.Id, Message = rule.Msg, Line = lines[0], Lines = sb.ToString(), Count = count, Snippet = snippet });
            }
        }

        // Finds constructs that the IE11 engine does not support. A page can switch rules off with a comment:
        //   <!-- ppt-html-allow: api-promise api-find -->   (e.g. after adding a polyfill)
        public static void Lint(string doc, List<Issue> issues, IEnumerable<string> extraAllow = null)
        {
            var allow = new HashSet<string>();
            if (extraAllow != null) foreach (string id in extraAllow) allow.Add(id);
            foreach (Match m in Regex.Matches(doc, @"ppt-html-allow\s*:\s*([\w\s,-]+)", IC))
                foreach (string id in m.Groups[1].Value.Split(new[] { ' ', ',', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries)) allow.Add(id.Trim().TrimEnd('-'));
            string body = Regex.Replace(doc, @"data:[\w/+.-]+;base64,[A-Za-z0-9+/=]+", delegate(Match m) { return new string(' ', m.Length); });
            // blank the harness blocks (keep length) so that they are not checked
            body = Regex.Replace(body, Regex.Escape(HeadOpen) + @"[\s\S]*?" + Regex.Escape(HeadClose) + "|" + Regex.Escape(TailOpen) + @"[\s\S]*?" + Regex.Escape(TailClose), delegate(Match m) { return Regex.Replace(m.Value, @"[^\n]", " "); });
            int pos = 0;
            var markup = new StringBuilder();
            foreach (Match m in SegRx.Matches(body))
            {
                markup.Append(body, pos, m.Index - pos);
                string t = m.Value;
                int line = LineOf(body, m.Index);
                if (t.StartsWith("<!--")) { markup.Append(Regex.Replace(t, @"[^\n]", " ")); }
                else
                {
                    int gt = t.IndexOf('>');
                    int lt = t.LastIndexOf("</", StringComparison.Ordinal);
                    string open = t.Substring(0, gt + 1), inner = t.Substring(gt + 1, lt - gt - 1);
                    int innerLine = line + LineOf(open, open.Length) - 1;
                    bool script = (t[1] == 's' || t[1] == 'S') && (t[2] == 'c' || t[2] == 'C');
                    if (script)
                    {
                        string type = (Attr(open, "type") ?? "").ToLowerInvariant();
                        bool js = type == "" || type.Contains("javascript") || type.Contains("ecmascript") || type == "module";
                        if (js) Apply(JsRules, CodeOnlyJs(inner), inner, innerLine, issues, allow);
                    }
                    else Apply(CssRules, BlankCssComments(inner), inner, innerLine, issues, allow);
                    markup.Append(open);
                    markup.Append(Regex.Replace(inner, @"[^\n]", " "));
                    markup.Append(t.Substring(lt));
                }
                pos = m.Index + m.Length;
            }
            markup.Append(body, pos, body.Length - pos);
            string mk = markup.ToString();
            Apply(HtmlRules, mk, mk, 1, issues, allow);
            // style attributes and inline event handlers
            foreach (Match m in Regex.Matches(mk, @"\sstyle\s*=\s*(""[^""]*""|'[^']*')", IC))
                Apply(CssRules, m.Groups[1].Value, m.Groups[1].Value, LineOf(mk, m.Index), issues, allow);
            foreach (Match m in Regex.Matches(mk, @"\son[a-z]+\s*=\s*""([^""]*)""", IC))
                Apply(JsRules, CodeOnlyJs(m.Groups[1].Value), m.Groups[1].Value, LineOf(mk, m.Index), issues, allow);
            issues.Sort(delegate(Issue a, Issue b) { int c = string.CompareOrdinal(a.Severity, b.Severity); return c != 0 ? c : a.Line.CompareTo(b.Line); });
        }

        // ------------------------------------------------------------------ helpers for the callers

        public static string XmlEscape(string s)
        {
            var sb = new StringBuilder(s.Length + s.Length / 8);
            foreach (char c in s)
            {
                if (c == '<') sb.Append("&lt;");
                else if (c == '>') sb.Append("&gt;");
                else if (c == '&') sb.Append("&amp;");
                else if (c == '\r') { }
                else if (c < 0x20 && c != '\n' && c != '\t') { }
                else sb.Append(c);
            }
            return sb.ToString();
        }

        // Text for a <script> string literal (JSON style), safe inside HTML.
        public static string JsString(string s)
        {
            var sb = new StringBuilder(s.Length + 64);
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '"': sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': break;
                    case '\t': sb.Append("\\t"); break;
                    case '<': sb.Append("\\u003C"); break;
                    case '>': sb.Append("\\u003E"); break;
                    case '&': sb.Append("\\u0026"); break;
                    case (char)0x2028: sb.Append("\\u2028"); break;
                    case (char)0x2029: sb.Append("\\u2029"); break;
                    default: if (c < 0x20) sb.Append("\\u" + ((int)c).ToString("X4")); else sb.Append(c); break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }

        public static void WriteUtf16(string path, string text)
        {
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                fs.WriteByte(0xFF); fs.WriteByte(0xFE);
                byte[] b = Encoding.Unicode.GetBytes(text);
                fs.Write(b, 0, b.Length);
            }
        }
    }
}
