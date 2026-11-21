// Reads what the build needs to know about font files (TrueType / OpenType / collections): family names,
// weight, licence flags, variable weight range. Only the few tables involved are read, not the whole file.
// C# 5 only (compiled by Windows PowerShell's Add-Type).
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace FontKit
{
    public class Face
    {
        public string File = "";
        public int Index;                 // face index inside a collection (.ttc)
        public string[] Families = new string[0];
        public int Weight = 400;
        public bool Italic;
        public int FsType;                // OS/2 embedding flags: 0 = installable embedding
        public bool Variable;             // has a weight axis
        public int WghtMin, WghtMax, WghtDefault;
        public bool Bundled;              // from the skill's own font pack
        public bool Embeddable { get { return (FsType & 0x010E) == 0; } }
    }

    public static class Fonts
    {
        static uint U32(byte[] b, int o) { return (uint)(b[o] << 24 | b[o + 1] << 16 | b[o + 2] << 8 | b[o + 3]); }
        static int U16(byte[] b, int o) { return b[o] << 8 | b[o + 1]; }

        static byte[] ReadAt(FileStream fs, long offset, int length)
        {
            if (offset < 0 || length <= 0 || offset + length > fs.Length) return null;
            var buf = new byte[length];
            fs.Position = offset;
            int got = 0;
            while (got < length) { int n = fs.Read(buf, got, length - got); if (n <= 0) return null; got += n; }
            return buf;
        }

        public static List<Face> Scan(string dir, bool bundled)
        {
            var list = new List<Face>();
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) return list;
            string[] files;
            try { files = Directory.GetFiles(dir); } catch { return list; }
            foreach (var f in files)
            {
                var ext = Path.GetExtension(f).ToLowerInvariant();
                if (ext != ".ttf" && ext != ".otf" && ext != ".ttc") continue;
                try { list.AddRange(ReadFile(f, bundled)); } catch { }
            }
            return list;
        }

        public static List<Face> ReadFile(string path, bool bundled)
        {
            var list = new List<Face>();
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                var head = ReadAt(fs, 0, 12);
                if (head == null) return list;
                if (U32(head, 0) == 0x74746366)       // 'ttcf'
                {
                    int n = (int)U32(head, 8);
                    var offs = ReadAt(fs, 12, 4 * n);
                    if (offs == null) return list;
                    for (int i = 0; i < n && i < 64; i++)
                    {
                        var face = ReadFace(fs, U32(offs, i * 4));
                        if (face != null) { face.File = path; face.Index = i; face.Bundled = bundled; list.Add(face); }
                    }
                }
                else
                {
                    var face = ReadFace(fs, 0);
                    if (face != null) { face.File = path; face.Bundled = bundled; list.Add(face); }
                }
            }
            return list;
        }

        static Face ReadFace(FileStream fs, long start)
        {
            var ot = ReadAt(fs, start, 12);
            if (ot == null) return null;
            uint ver = U32(ot, 0);
            if (ver != 0x00010000 && ver != 0x4F54544F && ver != 0x74727565) return null;   // TrueType, 'OTTO', 'true'
            int numTables = U16(ot, 4);
            var dir = ReadAt(fs, start + 12, 16 * numTables);
            if (dir == null) return null;
            long nameOff = 0, os2Off = 0, fvarOff = 0; int nameLen = 0, os2Len = 0, fvarLen = 0;
            for (int i = 0; i < numTables; i++)
            {
                uint tag = U32(dir, i * 16); long off = U32(dir, i * 16 + 8); int len = (int)U32(dir, i * 16 + 12);
                if (tag == 0x6E616D65) { nameOff = off; nameLen = len; }
                else if (tag == 0x4F532F32) { os2Off = off; os2Len = len; }
                else if (tag == 0x66766172) { fvarOff = off; fvarLen = len; }
            }
            var face = new Face();
            var fams = new List<string>();
            if (nameOff > 0 && nameLen > 6 && nameLen < 4 * 1024 * 1024)
            {
                var nt = ReadAt(fs, nameOff, nameLen);
                if (nt != null)
                {
                    int count = U16(nt, 2), strOff = U16(nt, 4);
                    for (int k = 0; k < count; k++)
                    {
                        int r = 6 + k * 12;
                        if (r + 12 > nt.Length) break;
                        int platform = U16(nt, r), nameId = U16(nt, r + 6), len = U16(nt, r + 8), so = strOff + U16(nt, r + 10);
                        if (nameId != 1 && nameId != 16) continue;
                        if (so + len > nt.Length || len == 0) continue;
                        string s = null;
                        if (platform == 3 || platform == 0) s = Encoding.BigEndianUnicode.GetString(nt, so, len);
                        else if (platform == 1 && U16(nt, r + 2) == 0) s = Encoding.ASCII.GetString(nt, so, len);
                        if (string.IsNullOrEmpty(s)) continue;
                        s = s.Trim();
                        bool dup = false;
                        foreach (var x in fams) if (string.Equals(x, s, StringComparison.OrdinalIgnoreCase)) { dup = true; break; }
                        if (!dup && s.Length > 0) fams.Add(s);
                    }
                }
            }
            if (fams.Count == 0) return null;
            // "Pretendard Variable" is asked for as "Pretendard" in style sheets
            foreach (var n0 in fams.ToArray())
            {
                foreach (var tail in new[] { " Variable", " VF" })
                {
                    if (!n0.EndsWith(tail, StringComparison.OrdinalIgnoreCase)) continue;
                    string n1 = n0.Substring(0, n0.Length - tail.Length).Trim();
                    bool dup = false;
                    foreach (var x in fams) if (string.Equals(x, n1, StringComparison.OrdinalIgnoreCase)) { dup = true; break; }
                    if (!dup && n1.Length > 0) fams.Add(n1);
                }
            }
            face.Families = fams.ToArray();
            if (os2Off > 0 && os2Len >= 64)
            {
                var os2 = ReadAt(fs, os2Off, 64);
                if (os2 != null) { face.Weight = U16(os2, 4); face.FsType = U16(os2, 8); face.Italic = (U16(os2, 62) & 1) != 0; }
            }
            if (face.Weight < 1 || face.Weight > 1000) face.Weight = 400;
            if (fvarOff > 0 && fvarLen >= 16 && fvarLen < 65536)
            {
                var fv = ReadAt(fs, fvarOff, fvarLen);
                if (fv != null)
                {
                    int axesOff = U16(fv, 4), axisCount = U16(fv, 8), axisSize = U16(fv, 10);
                    for (int a = 0; a < axisCount; a++)
                    {
                        int p = axesOff + a * axisSize;
                        if (p + 16 > fv.Length) break;
                        if (U32(fv, p) != 0x77676874) continue;      // 'wght'
                        face.Variable = true;
                        face.WghtMin = (int)Math.Round((int)U32(fv, p + 4) / 65536.0);
                        face.WghtDefault = (int)Math.Round((int)U32(fv, p + 8) / 65536.0);
                        face.WghtMax = (int)Math.Round((int)U32(fv, p + 12) / 65536.0);
                    }
                }
            }
            return face;
        }

        // Does the face have a glyph for this character (looked up in its cmap table, formats 4 and 12)?
        public static bool HasChar(string path, int faceIndex, int codePoint)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    long start = 0;
                    var head = ReadAt(fs, 0, 12);
                    if (head == null) return false;
                    if (U32(head, 0) == 0x74746366)
                    {
                        var offs = ReadAt(fs, 12 + 4 * faceIndex, 4);
                        if (offs == null) return false;
                        start = U32(offs, 0);
                    }
                    var ot = ReadAt(fs, start, 12);
                    if (ot == null) return false;
                    int numTables = U16(ot, 4);
                    var dir = ReadAt(fs, start + 12, 16 * numTables);
                    if (dir == null) return false;
                    long cmapOff = 0; int cmapLen = 0;
                    for (int i = 0; i < numTables; i++) if (U32(dir, i * 16) == 0x636D6170) { cmapOff = U32(dir, i * 16 + 8); cmapLen = (int)U32(dir, i * 16 + 12); }
                    if (cmapOff == 0 || cmapLen < 4 || cmapLen > 16 * 1024 * 1024) return false;
                    var cm = ReadAt(fs, cmapOff, cmapLen);
                    if (cm == null) return false;
                    int n = U16(cm, 2);
                    for (int i = 0; i < n; i++)
                    {
                        int rec = 4 + i * 8;
                        if (rec + 8 > cm.Length) break;
                        int sub = (int)U32(cm, rec + 4);
                        if (sub + 8 > cm.Length) continue;
                        int format = U16(cm, sub);
                        if (format == 12 && sub + 16 <= cm.Length)
                        {
                            int groups = (int)U32(cm, sub + 12);
                            for (int g = 0; g < groups; g++)
                            {
                                int p = sub + 16 + g * 12;
                                if (p + 12 > cm.Length) break;
                                uint a = U32(cm, p), b = U32(cm, p + 4);
                                if ((uint)codePoint >= a && (uint)codePoint <= b) return true;
                            }
                        }
                        else if (format == 4 && codePoint <= 0xFFFF)
                        {
                            int segX2 = U16(cm, sub + 6), endAt = sub + 14, startAt = endAt + segX2 + 2;
                            for (int sgm = 0; sgm < segX2 / 2; sgm++)
                            {
                                if (startAt + sgm * 2 + 2 > cm.Length) break;
                                int end = U16(cm, endAt + sgm * 2), st = U16(cm, startAt + sgm * 2);
                                if (codePoint >= st && codePoint <= end && end != 0xFFFF) return true;
                            }
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        // Distinct characters of a text (as one string), ignoring base64 payloads; used as the subset of a page.
        public static string Characters(string text)
        {
            var seen = new HashSet<int>();
            var sb = new StringBuilder();
            int i = 0, n = text.Length;
            while (i < n)
            {
                // skip "base64,XXXX" runs quickly
                if (text[i] == 'b' && i + 7 < n && string.CompareOrdinal(text, i, "base64,", 0, 7) == 0)
                {
                    i += 7;
                    while (i < n) { char c = text[i]; if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '+' || c == '/' || c == '=') i++; else break; }
                    continue;
                }
                int cp; int step = 1;
                if (char.IsHighSurrogate(text[i]) && i + 1 < n && char.IsLowSurrogate(text[i + 1])) { cp = char.ConvertToUtf32(text[i], text[i + 1]); step = 2; }
                else cp = text[i];
                if (cp >= 32 && seen.Add(cp)) sb.Append(text, i, step);
                i += step;
            }
            foreach (char c in "0123456789.,:;!?%+-()[]/'\"&*#@~<>=_ ") if (seen.Add(c)) sb.Append(c);
            return sb.ToString();
        }
    }
}
