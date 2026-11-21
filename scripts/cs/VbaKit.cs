// VBA project writer: compound file (OLE) container + MS-OVBA compression + project streams.
// Builds a source-only vbaProject.bin from text sources; Office compiles it when the file is opened.
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using System.Text;
using System;
namespace OleKit
{
    [ComImport, Guid("0000000d-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumSTATSTG
    {
        [PreserveSig]
        int Next(uint celt, [MarshalAs(UnmanagedType.LPArray), Out] System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt, out uint pceltFetched);
        void Skip(uint celt);
        void Reset();
        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumSTATSTG Clone();
    }

    [ComImport, Guid("0000000b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStorage
    {
        void CreateStream([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, uint grfMode, uint reserved1, uint reserved2, out IStream ppstm);
        void OpenStream([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr reserved1, uint grfMode, uint reserved2, out IStream ppstm);
        void CreateStorage([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, uint grfMode, uint reserved1, uint reserved2, out IStorage ppstg);
        void OpenStorage([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IStorage pstgPriority, uint grfMode, IntPtr snbExclude, uint reserved, out IStorage ppstg);
        void CopyTo(uint ciidExclude, IntPtr rgiidExclude, IntPtr snbExclude, IStorage pstgDest);
        void MoveElementTo([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IStorage pstgDest, [MarshalAs(UnmanagedType.LPWStr)] string pwcsNewName, uint grfFlags);
        void Commit(uint grfCommitFlags);
        void Revert();
        void EnumElements(uint reserved1, IntPtr reserved2, uint reserved3, out IEnumSTATSTG ppenum);
        void DestroyElement([MarshalAs(UnmanagedType.LPWStr)] string pwcsName);
        void RenameElement([MarshalAs(UnmanagedType.LPWStr)] string pwcsOldName, [MarshalAs(UnmanagedType.LPWStr)] string pwcsNewName);
        void SetElementTimes([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr pctime, IntPtr patime, IntPtr pmtime);
        void SetClass(ref Guid clsid);
        void SetStateBits(uint grfStateBits, uint grfMask);
        void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, uint grfStatFlag);
    }

    public class Entry
    {
        public string Path;      // e.g. "VBA/dir"
        public bool IsStorage;
        public long Size;
        public Guid Clsid;
        public byte[] Data;      // streams only
    }

    public static class Cfb
    {
        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        static extern int StgCreateDocfile(string pwcsName, uint grfMode, uint reserved, out IStorage ppstgOpen);
        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        static extern int StgOpenStorage(string pwcsName, IStorage pstgPriority, uint grfMode, IntPtr snbExclude, uint reserved, out IStorage ppstgOpen);

        const uint STGM_READ = 0x0, STGM_READWRITE = 0x2, STGM_SHARE_EXCLUSIVE = 0x10, STGM_SHARE_DENY_WRITE = 0x20, STGM_CREATE = 0x1000;

        public static List<Entry> ReadAll(string file)
        {
            IStorage root;
            int hr = StgOpenStorage(file, null, STGM_READ | STGM_SHARE_DENY_WRITE, IntPtr.Zero, 0, out root);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            var list = new List<Entry>();
            try
            {
                System.Runtime.InteropServices.ComTypes.STATSTG st;
                root.Stat(out st, 1);
                list.Add(new Entry { Path = "", IsStorage = true, Clsid = st.clsid });
                Walk(root, "", list);
            }
            finally { Marshal.ReleaseComObject(root); }
            return list;
        }

        static void Walk(IStorage stg, string prefix, List<Entry> list)
        {
            IEnumSTATSTG en;
            stg.EnumElements(0, IntPtr.Zero, 0, out en);
            var arr = new System.Runtime.InteropServices.ComTypes.STATSTG[1];
            uint got;
            while (en.Next(1, arr, out got) == 0 && got == 1)
            {
                var s = arr[0];
                string p = prefix + s.pwcsName;
                if (s.type == 1)
                {
                    IStorage child;
                    stg.OpenStorage(s.pwcsName, null, STGM_READ | STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out child);
                    list.Add(new Entry { Path = p, IsStorage = true, Clsid = s.clsid });
                    Walk(child, p + "/", list);
                    Marshal.ReleaseComObject(child);
                }
                else if (s.type == 2)
                {
                    IStream sm;
                    stg.OpenStream(s.pwcsName, IntPtr.Zero, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, out sm);
                    byte[] buf = new byte[(int)s.cbSize];
                    if (buf.Length > 0)
                    {
                        IntPtr pcb = Marshal.AllocHGlobal(8);
                        try { sm.Read(buf, buf.Length, pcb); }
                        finally { Marshal.FreeHGlobal(pcb); }
                    }
                    Marshal.ReleaseComObject(sm);
                    list.Add(new Entry { Path = p, IsStorage = false, Size = s.cbSize, Data = buf });
                }
            }
            Marshal.ReleaseComObject(en);
        }

        // entries: storages must be listed before their children. Path separator '/'.
        public static void Write(string file, List<Entry> entries)
        {
            IStorage root;
            int hr = StgCreateDocfile(file, STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, out root);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            var open = new Dictionary<string, IStorage>();
            open[""] = root;
            try
            {
                foreach (var e in entries)
                {
                    if (e.Path == "")
                    {
                        if (e.Clsid != Guid.Empty) { Guid g = e.Clsid; root.SetClass(ref g); }
                        continue;
                    }
                    int ix = e.Path.LastIndexOf('/');
                    string parent = ix < 0 ? "" : e.Path.Substring(0, ix);
                    string name = ix < 0 ? e.Path : e.Path.Substring(ix + 1);
                    IStorage ps = open[parent];
                    if (e.IsStorage)
                    {
                        IStorage child;
                        ps.CreateStorage(name, STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, out child);
                        if (e.Clsid != Guid.Empty) { Guid g = e.Clsid; child.SetClass(ref g); }
                        open[e.Path] = child;
                    }
                    else
                    {
                        IStream sm;
                        ps.CreateStream(name, STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, out sm);
                        byte[] d = e.Data ?? new byte[0];
                        if (d.Length > 0) sm.Write(d, d.Length, IntPtr.Zero);
                        sm.Commit(0);
                        Marshal.ReleaseComObject(sm);
                    }
                }
                // commit children first (deepest paths first)
                var keys = new List<string>(open.Keys);
                keys.Sort(delegate(string a, string b) { return b.Length.CompareTo(a.Length); });
                foreach (var k in keys) { open[k].Commit(0); }
            }
            finally
            {
                var keys = new List<string>(open.Keys);
                keys.Sort(delegate(string a, string b) { return b.Length.CompareTo(a.Length); });
                foreach (var k in keys) { Marshal.ReleaseComObject(open[k]); }
            }
        }
    }

    public static class Ovba
    {
        public static byte[] Decompress(byte[] src, int offset)
        {
            var o = new MemoryStream();
            int p = offset;
            if (src[p] != 1) throw new InvalidDataException("bad signature byte " + src[p]);
            p++;
            while (p + 1 < src.Length)
            {
                int hdr = src[p] | (src[p + 1] << 8);
                int start = p;
                p += 2;
                int size = (hdr & 0x0FFF) + 3;
                bool comp = (hdr & 0x8000) != 0;
                int end = Math.Min(src.Length, start + size);
                if (!comp)
                {
                    int n0 = Math.Min(4096, src.Length - p);
                    o.Write(src, p, n0);
                    p += 4096;
                    continue;
                }
                byte[] buf = new byte[4096 + 300];
                int n = 0;
                while (p < end)
                {
                    int flags = src[p++];
                    for (int bit = 0; bit < 8 && p < end; bit++)
                    {
                        if ((flags & (1 << bit)) == 0) { buf[n++] = src[p++]; }
                        else
                        {
                            int tok = src[p] | (src[p + 1] << 8);
                            p += 2;
                            int bc = 4;
                            while ((1 << bc) < n) bc++;
                            int lenMask = 0xFFFF >> bc;
                            int len = (tok & lenMask) + 3;
                            int off = (tok >> (16 - bc)) + 1;
                            for (int k = 0; k < len; k++) { buf[n] = buf[n - off]; n++; }
                        }
                    }
                }
                o.Write(buf, 0, n);
            }
            return o.ToArray();
        }

        public static byte[] Compress(byte[] data)
        {
            var o = new MemoryStream();
            o.WriteByte(1);
            for (int cs = 0; cs < data.Length; cs += 4096)
            {
                int n = Math.Min(4096, data.Length - cs);
                byte[] c = CompressChunk(data, cs, n);
                if (c.Length > 4096)
                {
                    int hdr = 0x3000 | 0x0FFF;
                    o.WriteByte((byte)(hdr & 0xFF)); o.WriteByte((byte)(hdr >> 8));
                    byte[] raw = new byte[4096];
                    Array.Copy(data, cs, raw, 0, n);
                    o.Write(raw, 0, 4096);
                }
                else
                {
                    int hdr = 0xB000 | (c.Length + 2 - 3);
                    o.WriteByte((byte)(hdr & 0xFF)); o.WriteByte((byte)(hdr >> 8));
                    o.Write(c, 0, c.Length);
                }
            }
            return o.ToArray();
        }

        static byte[] CompressChunk(byte[] d, int cs, int n)
        {
            var o = new List<byte>(n + n / 8 + 4);
            int pos = 0;
            while (pos < n)
            {
                int flagIdx = o.Count;
                o.Add(0);
                int flags = 0;
                for (int bit = 0; bit < 8 && pos < n; bit++)
                {
                    int bestLen = 0, bestOff = 0;
                    if (pos > 0)
                    {
                        int bc = 4;
                        while ((1 << bc) < pos) bc++;
                        int maxLen = (0xFFFF >> bc) + 3;
                        for (int cand = pos - 1; cand >= 0; cand--)
                        {
                            int len = 0;
                            while (pos + len < n && len < maxLen && d[cs + cand + len] == d[cs + pos + len]) len++;
                            if (len > bestLen) { bestLen = len; bestOff = pos - cand; if (len == maxLen) break; }
                        }
                        if (bestLen >= 3)
                        {
                            int tok = ((bestOff - 1) << (16 - bc)) | (bestLen - 3);
                            o.Add((byte)(tok & 0xFF)); o.Add((byte)((tok >> 8) & 0xFF));
                            flags |= (1 << bit);
                            pos += bestLen;
                            continue;
                        }
                    }
                    o.Add(d[cs + pos]);
                    pos++;
                }
                o[flagIdx] = (byte)flags;
            }
            return o.ToArray();
        }
    }
}

namespace OleKit
{
    // Builds a source-only vbaProject.bin (MS-OVBA) with standard modules, class modules and one
    // empty MSForms UserForm (designer storage per MS-OFORMS). Office recompiles it on first load.
    public static class VbaProj
    {
        static readonly Encoding Ascii = Encoding.ASCII;
        static readonly byte[] Empty = new byte[0];

        static byte[] A(string s) { return Ascii.GetBytes(s); }
        static byte[] Un(string s) { return Encoding.Unicode.GetBytes(s); }

        static void Rec(BinaryWriter w, ushort id, byte[] data) { w.Write(id); w.Write((uint)data.Length); w.Write(data); }
        static void RecU32(BinaryWriter w, ushort id, uint v) { w.Write(id); w.Write((uint)4); w.Write(v); }
        static void RecU16(BinaryWriter w, ushort id, ushort v) { w.Write(id); w.Write((uint)2); w.Write(v); }

        static void Reference(BinaryWriter w, string name, string libid)
        {
            Rec(w, 0x0016, A(name));
            Rec(w, 0x003E, Un(name));
            byte[] l = A(libid);
            w.Write((ushort)0x000D);
            w.Write((uint)(4 + l.Length + 6));
            w.Write((uint)l.Length);
            w.Write(l);
            w.Write((uint)0);
            w.Write((ushort)0);
        }

        // MS-OVBA 2.4.3.2 data encryption (used for CMG / DPB / GC in the PROJECT stream)
        static string Encrypt(byte[] data, string projectId, byte seed)
        {
            var o = new List<byte>();
            int projKey = 0;
            foreach (char c in projectId) projKey = (projKey + (byte)c) & 0xFF;
            byte version = 2;
            byte versionEnc = (byte)(seed ^ version);
            byte projKeyEnc = (byte)(seed ^ projKey);
            byte unenc1 = (byte)projKey, enc1 = projKeyEnc, enc2 = versionEnc;
            o.Add(seed); o.Add(versionEnc); o.Add(projKeyEnc);
            int ignored = (seed & 6) / 2;
            var all = new List<byte>();
            for (int i = 0; i < ignored; i++) all.Add((byte)(0x51 + i * 7));
            int nIgnored = all.Count;
            all.AddRange(BitConverter.GetBytes((uint)data.Length));
            all.AddRange(data);
            foreach (byte b in all)
            {
                byte e = (byte)(b ^ ((enc2 + unenc1) & 0xFF));
                o.Add(e);
                enc2 = enc1; enc1 = e; unenc1 = b;
            }
            var sb = new StringBuilder();
            foreach (byte b in o) sb.Append(b.ToString("X2"));
            return sb.ToString();
        }

        static string ClassHeader(string name, string vbBase, bool predeclared)
        {
            return "Attribute VB_Name = \"" + name + "\"\r\n" +
                   "Attribute VB_Base = \"" + vbBase + "\"\r\n" +
                   "Attribute VB_GlobalNameSpace = False\r\n" +
                   "Attribute VB_Creatable = False\r\n" +
                   "Attribute VB_PredeclaredId = " + (predeclared ? "True" : "False") + "\r\n" +
                   "Attribute VB_Exposed = False\r\n" +
                   "Attribute VB_TemplateDerived = False\r\n" +
                   "Attribute VB_Customizable = False\r\n";
        }

        static byte[] FormStreamF(int widthPt, int heightPt, uint shapeCookie)
        {
            var ms = new MemoryStream();
            var w = new BinaryWriter(ms);
            w.Write((byte)0x00); w.Write((byte)0x04);          // MinorVersion, MajorVersion
            w.Write((ushort)0x0028);                           // cbForm
            w.Write((uint)0x0C100C48);                         // PropMask: NextAvailableID, BooleanProperties, DisplayedSize, LogicalSize, Font, ShapeCookie, DrawBuffer
            w.Write((uint)1);                                  // NextAvailableID
            w.Write((uint)0x00004004);                         // BooleanProperties: ENABLED | DESINKPERSISTED
            w.Write((ushort)0xFFFF); w.Write((ushort)0);       // Font marker + padding
            w.Write(shapeCookie);                              // ShapeCookie
            w.Write((uint)32000);                              // DrawBuffer
            w.Write((uint)Math.Round(widthPt * 2540.0 / 72));  // DisplayedSize (HIMETRIC)
            w.Write((uint)Math.Round(heightPt * 2540.0 / 72));
            w.Write((uint)0); w.Write((uint)0);                // LogicalSize
            w.Write(new Guid("0BE35203-8F91-11CE-9DE3-00AA004BB851").ToByteArray()); // StdFont
            w.Write((byte)1);                                  // StdFont version
            w.Write((ushort)0);                                // charset
            w.Write((byte)0);                                  // flags
            w.Write((ushort)400);                              // weight
            w.Write((uint)82500);                              // height (8.25pt)
            w.Write((byte)6); w.Write(A("Tahoma"));
            w.Write((ushort)0);                                // CountOfSiteClassInfo
            w.Write((uint)0);                                  // CountOfSites
            w.Write((uint)0);                                  // CountOfBytes
            w.Write(new byte[] { 0x00, 0x02, 0x0C, 0x00, 0x19, 0x00, 0x00, 0x00, 0xF3, 0x7F, 0x01, 0x00, 0xFF, 0x01, 0x00, 0x00 }); // DesignExtender
            w.Flush();
            return ms.ToArray();
        }

        static byte[] FormCompObj()
        {
            var ms = new MemoryStream();
            var w = new BinaryWriter(ms);
            w.Write(new byte[] { 0x01, 0x00, 0xFE, 0xFF, 0x03, 0x0A, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF });
            w.Write(new byte[16]);
            byte[] a = A("Microsoft Forms 2.0 Form\0"); w.Write((uint)a.Length); w.Write(a);
            byte[] b = A("Embedded Object\0"); w.Write((uint)b.Length); w.Write(b);
            w.Write((uint)0);
            w.Write(new byte[] { 0xF4, 0x39, 0xB2, 0x71 });
            w.Write(new byte[12]);
            w.Flush();
            return ms.ToArray();
        }

        // moduleFiles: file names inside srcDir; extension decides the kind (.bas module, .cls class, .frm UserForm).
        // refs: "name|libid" strings.
        public static string Build(string outFile, string srcDir, string[] moduleFiles, string projectName, string[] refs, int codePage, uint lcid, uint sysKind, int formWidthPt, int formHeightPt, string formCaption)
        {
            var log = new StringBuilder();
            var names = new List<string>();
            var kinds = new List<int>();   // 0 std, 1 class, 2 form
            var sources = new List<byte[]>();
            string formName = null;
            foreach (string f in moduleFiles)
            {
                string name = Path.GetFileNameWithoutExtension(f);
                string ext = Path.GetExtension(f).ToLowerInvariant();
                string code = File.ReadAllText(Path.Combine(srcDir, f), Ascii).Replace("\r\n", "\n").Replace("\n", "\r\n");
                foreach (char c in code) if (c > 127) throw new InvalidDataException(f + ": non-ASCII source");
                int kind = ext == ".bas" ? 0 : (ext == ".cls" ? 1 : 2);
                string header;
                if (kind == 0) header = "Attribute VB_Name = \"" + name + "\"\r\n";
                else if (kind == 1) header = ClassHeader(name, "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}", false);
                else
                {
                    header = ClassHeader(name, "0{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}", true);
                    if (formName != null) throw new InvalidDataException("only one form supported");
                    formName = name;
                }
                names.Add(name); kinds.Add(kind); sources.Add(A(header + code));
                log.AppendLine("module " + name + " kind=" + kind + " source=" + (header.Length + code.Length) + " chars");
            }

            // ---- dir stream
            var dms = new MemoryStream();
            var w = new BinaryWriter(dms);
            RecU32(w, 0x0001, sysKind);
            RecU32(w, 0x0002, lcid);
            RecU32(w, 0x0014, 0x0409);
            RecU16(w, 0x0003, (ushort)codePage);
            Rec(w, 0x0004, A(projectName));
            Rec(w, 0x0005, Empty); Rec(w, 0x0040, Empty);
            Rec(w, 0x0006, Empty); Rec(w, 0x003D, Empty);
            RecU32(w, 0x0007, 0);
            RecU32(w, 0x0008, 0);
            w.Write((ushort)0x0009); w.Write((uint)4); w.Write((uint)0x6532A1B4); w.Write((ushort)0x000C);
            Rec(w, 0x000C, Empty); Rec(w, 0x003C, Empty);
            foreach (string r in refs)
            {
                int ix = r.IndexOf('|');
                Reference(w, r.Substring(0, ix), r.Substring(ix + 1));
            }
            RecU16(w, 0x000F, (ushort)names.Count);
            RecU16(w, 0x0013, 0xFFFF);
            for (int i = 0; i < names.Count; i++)
            {
                Rec(w, 0x0019, A(names[i]));
                Rec(w, 0x0047, Un(names[i]));
                Rec(w, 0x001A, A(names[i]));
                Rec(w, 0x0032, Un(names[i]));
                Rec(w, 0x001C, Empty);
                Rec(w, 0x0048, Empty);
                RecU32(w, 0x0031, 0);
                RecU32(w, 0x001E, 0);
                RecU16(w, 0x002C, 0xFFFF);
                Rec(w, (ushort)(kinds[i] == 0 ? 0x0021 : 0x0022), Empty);
                if (kinds[i] != 0) Rec(w, 0x0028, Empty);
                Rec(w, 0x002B, Empty);
            }
            w.Write((ushort)0x0010); w.Write((uint)0);
            w.Flush();
            byte[] dirRaw = dms.ToArray();
            log.AppendLine("dir raw=" + dirRaw.Length);

            // ---- PROJECT stream
            string id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
            var p = new StringBuilder();
            p.Append("ID=\"" + id + "\"\r\n");
            for (int i = 0; i < names.Count; i++)
            {
                if (kinds[i] == 0) p.Append("Module=" + names[i] + "\r\n");
                else if (kinds[i] == 1) p.Append("Class=" + names[i] + "\r\n");
            }
            if (formName != null)
            {
                p.Append("Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}\r\n");
                p.Append("BaseClass=" + formName + "\r\n");
            }
            p.Append("Name=\"" + projectName + "\"\r\n");
            p.Append("HelpContextID=\"0\"\r\n");
            p.Append("VersionCompatible32=\"393222000\"\r\n");
            p.Append("CMG=\"" + Encrypt(new byte[] { 0, 0, 0, 0 }, id, 0x3A) + "\"\r\n");
            p.Append("DPB=\"" + Encrypt(new byte[] { 0 }, id, 0x74) + "\"\r\n");
            p.Append("GC=\"" + Encrypt(new byte[] { 0xFF }, id, 0xAE) + "\"\r\n");
            p.Append("\r\n[Host Extender Info]\r\n&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n");
            p.Append("\r\n[Workspace]\r\n");
            for (int i = 0; i < names.Count; i++)
                p.Append(names[i] + (kinds[i] == 2 ? "=0, 0, 0, 0, C, 26, 26, 1200, 700, C\r\n" : "=0, 0, 0, 0, C\r\n"));

            // ---- PROJECTwm
            var wm = new MemoryStream();
            foreach (string n in names) { byte[] a = A(n); wm.Write(a, 0, a.Length); wm.WriteByte(0); byte[] u = Un(n); wm.Write(u, 0, u.Length); wm.WriteByte(0); wm.WriteByte(0); }
            wm.WriteByte(0); wm.WriteByte(0);

            // ---- compound file
            var entries = new List<Entry>();
            entries.Add(new Entry { Path = "", IsStorage = true });
            entries.Add(new Entry { Path = "VBA", IsStorage = true });
            entries.Add(new Entry { Path = "VBA/dir", Data = Ovba.Compress(dirRaw) });
            entries.Add(new Entry { Path = "VBA/_VBA_PROJECT", Data = new byte[] { 0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00 } });
            for (int i = 0; i < names.Count; i++)
            {
                byte[] comp = Ovba.Compress(sources[i]);
                byte[] back = Ovba.Decompress(comp, 0);
                if (back.Length != sources[i].Length) throw new InvalidDataException("round-trip length mismatch for " + names[i]);
                for (int k = 0; k < back.Length; k++) if (back[k] != sources[i][k]) throw new InvalidDataException("round-trip mismatch for " + names[i] + " at " + k);
                entries.Add(new Entry { Path = "VBA/" + names[i], Data = comp });
                log.AppendLine("stream VBA/" + names[i] + " raw=" + sources[i].Length + " compressed=" + comp.Length);
            }
            entries.Add(new Entry { Path = "PROJECT", Data = A(p.ToString()) });
            entries.Add(new Entry { Path = "PROJECTwm", Data = wm.ToArray() });
            if (formName != null)
            {
                uint cookie = 2;
                string vbFrame = "VERSION 5.00\r\n" +
                    "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " + formName + " \r\n" +
                    "   Caption         =   \"" + formCaption + "\"\r\n" +
                    "   ClientHeight    =   " + (formHeightPt * 20) + "\r\n" +
                    "   ClientLeft      =   45\r\n" +
                    "   ClientTop       =   330\r\n" +
                    "   ClientWidth     =   " + (formWidthPt * 20) + "\r\n" +
                    "   StartUpPosition =   1  'CenterOwner\r\n" +
                    "   TypeInfoVer     =   " + cookie + "\r\n" +
                    "End\r\n";
                entries.Add(new Entry { Path = formName, IsStorage = true });
                entries.Add(new Entry { Path = formName + "/f", Data = FormStreamF(formWidthPt, formHeightPt, cookie) });
                entries.Add(new Entry { Path = formName + "/o", Data = Empty });
                entries.Add(new Entry { Path = formName + "/\u0001CompObj", Data = FormCompObj() });
                entries.Add(new Entry { Path = formName + "/\u0003VBFrame", Data = A(vbFrame) });
            }
            if (File.Exists(outFile)) File.Delete(outFile);
            Cfb.Write(outFile, entries);
            log.AppendLine("written " + outFile + " (" + new FileInfo(outFile).Length + " bytes)");
            return log.ToString();
        }
    }
}
