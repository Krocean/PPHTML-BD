using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;

namespace DeskKit
{
    public class WinInfo
    {
        public IntPtr Hwnd;
        public int Pid;
        public string Class;
        public string Title;
        public bool Visible;
        public int Left, Top, Width, Height;
        public string ChildText;
        public override string ToString() { return "hwnd=" + Hwnd + " pid=" + Pid + " class=" + Class + " vis=" + Visible + " title=[" + Title + "] " + Width + "x" + Height + " @" + Left + "," + Top; }
    }

    // Window helpers that can run against a separate (hidden) desktop: every call takes the desktop name
    // ("" = the caller's own desktop) and executes on a short-lived thread attached to that desktop.
    public static class Desk
    {
        delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)] static extern IntPtr CreateDesktop(string name, IntPtr device, IntPtr devmode, int flags, uint access, IntPtr sa);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)] static extern IntPtr OpenDesktop(string name, int flags, bool inherit, uint access);
        [DllImport("user32.dll")] static extern bool CloseDesktop(IntPtr h);
        [DllImport("user32.dll", SetLastError = true)] static extern bool SetThreadDesktop(IntPtr h);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr lParam);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder sb, int max);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder sb, int max);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT r);
        [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hwnd, out RECT r);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
        [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr hwnd);
        [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr hwnd, uint cmd);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] static extern IntPtr GetWindowLongPtr(IntPtr hwnd, int index);
        [DllImport("user32.dll")] static extern uint MapVirtualKey(uint code, uint mapType);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)] static extern bool CreateProcess(string app, StringBuilder cmd, IntPtr pa, IntPtr ta, bool inherit, uint flags, IntPtr env, string cwd, ref STARTUPINFO si, out PROCESS_INFORMATION pi);
        [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
        [StructLayout(LayoutKind.Sequential)] struct RECT { public int L, T, R, B; }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct STARTUPINFO
        {
            public int cb; public string lpReserved; public string lpDesktop; public string lpTitle;
            public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
            public short wShowWindow, cbReserved2; public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
        }
        [StructLayout(LayoutKind.Sequential)] struct PROCESS_INFORMATION { public IntPtr hProcess, hThread; public int dwProcessId, dwThreadId; }
        const uint GENERIC_ALL = 0x10000000;

        static string Text(IntPtr h) { var sb = new StringBuilder(1024); GetWindowText(h, sb, 1024); return sb.ToString(); }
        static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }

        [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr ctx);

        static T On<T>(string desk, Func<T> f)
        {
            T result = default(T); Exception err = null;
            var th = new Thread(delegate()
            {
                IntPtr h = IntPtr.Zero;
                try
                {
                    if (!string.IsNullOrEmpty(desk))
                    {
                        h = OpenDesktop(desk, 0, false, GENERIC_ALL);
                        if (h == IntPtr.Zero) throw new InvalidOperationException("OpenDesktop failed " + Marshal.GetLastWin32Error());
                        if (!SetThreadDesktop(h)) throw new InvalidOperationException("SetThreadDesktop failed " + Marshal.GetLastWin32Error());
                    }
                    // physical pixels (per-monitor aware), the same coordinates PowerPoint itself works in
                    try { SetThreadDpiAwarenessContext(new IntPtr(-4)); } catch { }
                    result = f();
                }
                catch (Exception e) { err = e; }
                finally { if (h != IntPtr.Zero) CloseDesktop(h); }
            });
            th.SetApartmentState(ApartmentState.MTA);
            th.Start(); th.Join();
            if (err != null) throw new InvalidOperationException(err.Message, err);
            return result;
        }

        public static bool Exists(string desk)
        {
            IntPtr h = OpenDesktop(desk, 0, false, GENERIC_ALL);
            if (h == IntPtr.Zero) return false;
            CloseDesktop(h); return true;
        }

        // Starts a process on the named desktop (created when missing). Returns the process id.
        public static int Start(string desk, string commandLine, string cwd)
        {
            IntPtr h = OpenDesktop(desk, 0, false, GENERIC_ALL);
            if (h == IntPtr.Zero) h = CreateDesktop(desk, IntPtr.Zero, IntPtr.Zero, 0, GENERIC_ALL, IntPtr.Zero);
            if (h == IntPtr.Zero) throw new InvalidOperationException("CreateDesktop failed " + Marshal.GetLastWin32Error());
            var si = new STARTUPINFO(); si.cb = Marshal.SizeOf(typeof(STARTUPINFO)); si.lpDesktop = desk;
            PROCESS_INFORMATION pi;
            var cmd = new StringBuilder(commandLine, 4096);
            if (!CreateProcess(null, cmd, IntPtr.Zero, IntPtr.Zero, false, 0, IntPtr.Zero, cwd, ref si, out pi))
            {
                int e = Marshal.GetLastWin32Error(); CloseDesktop(h);
                throw new InvalidOperationException("CreateProcess failed " + e);
            }
            CloseHandle(pi.hThread); CloseHandle(pi.hProcess);
            // the desktop handle is intentionally kept open for the life of this (short) helper process
            return pi.dwProcessId;
        }

        static WinInfo Info(IntPtr h, bool withChildText)
        {
            uint p; GetWindowThreadProcessId(h, out p);
            RECT r; GetWindowRect(h, out r);
            var wi = new WinInfo { Hwnd = h, Pid = (int)p, Class = Cls(h), Title = Text(h), Visible = IsWindowVisible(h), Left = r.L, Top = r.T, Width = r.R - r.L, Height = r.B - r.T, ChildText = "" };
            if (withChildText)
            {
                var sb = new StringBuilder();
                EnumChildWindows(h, delegate(IntPtr c, IntPtr l2)
                {
                    string t = Text(c);
                    if (!string.IsNullOrEmpty(t)) sb.Append("[" + Cls(c) + ": " + t.Replace("\r", " ").Replace("\n", " ") + "] ");
                    return true;
                }, IntPtr.Zero);
                wi.ChildText = sb.ToString();
            }
            return wi;
        }

        // Top-level windows on the desktop; pid 0 = all processes.
        public static List<WinInfo> Windows(string desk, int pid)
        {
            return On(desk, delegate()
            {
                var list = new List<WinInfo>();
                EnumWindows(delegate(IntPtr h, IntPtr l)
                {
                    uint p; GetWindowThreadProcessId(h, out p);
                    if (pid == 0 || p == (uint)pid) list.Add(Info(h, true));
                    return true;
                }, IntPtr.Zero);
                return list;
            });
        }

        static string DescribeCore(IntPtr hwnd)
        {
            RECT r; GetWindowRect(hwnd, out r);
            RECT c; GetClientRect(hwnd, out c);
            long style = GetWindowLongPtr(hwnd, -16).ToInt64(), ex = GetWindowLongPtr(hwnd, -20).ToInt64();
            return "hwnd=" + hwnd + " class=" + Cls(hwnd) + " title=[" + Text(hwnd) + "] rect=" + r.L + "," + r.T + " " + (r.R - r.L) + "x" + (r.B - r.T) + " client=" + c.R + "x" + c.B + " vis=" + IsWindowVisible(hwnd) + " style=0x" + style.ToString("X") + " ex=0x" + ex.ToString("X") + " parent=" + GetParent(hwnd) + " owner=" + GetWindow(hwnd, 4);
        }
        public static string Describe(string desk, IntPtr hwnd) { return On(desk, delegate() { return IsWindow(hwnd) ? DescribeCore(hwnd) : "hwnd=" + hwnd + " (not a window)"; }); }
        public static List<string> Tree(string desk, IntPtr parent)
        {
            return On(desk, delegate()
            {
                var list = new List<string>();
                EnumChildWindows(parent, delegate(IntPtr c, IntPtr l) { list.Add(DescribeCore(c)); return true; }, IntPtr.Zero);
                return list;
            });
        }
        public static bool Alive(string desk, IntPtr hwnd) { return On(desk, delegate() { return IsWindow(hwnd); }); }
        public static IntPtr FindDescendant(string desk, IntPtr parent, string className, int index)
        {
            return On(desk, delegate()
            {
                IntPtr found = IntPtr.Zero; int n = 0;
                EnumChildWindows(parent, delegate(IntPtr c, IntPtr l)
                {
                    if (Cls(c) == className) { if (n == index) { found = c; return false; } n++; }
                    return true;
                }, IntPtr.Zero);
                return found;
            });
        }
        public static bool Close(string desk, IntPtr hwnd) { return On(desk, delegate() { return PostMessage(hwnd, 0x0010, IntPtr.Zero, IntPtr.Zero); }); }
        // Presses the n-th push button of a dialog (BM_CLICK); index 0 is normally the default button.
        public static bool ClickButton(string desk, IntPtr dialog, int index)
        {
            return On(desk, delegate()
            {
                IntPtr found = IntPtr.Zero; int n = 0;
                EnumChildWindows(dialog, delegate(IntPtr c, IntPtr l)
                {
                    if (Cls(c) == "Button") { if (n == index) { found = c; return false; } n++; }
                    return true;
                }, IntPtr.Zero);
                if (found == IntPtr.Zero) return false;
                return PostMessage(found, 0x00F5, IntPtr.Zero, IntPtr.Zero);
            });
        }

        public static bool Capture(string desk, IntPtr hwnd, string pngPath, uint flags)
        {
            return On(desk, delegate()
            {
                RECT r; GetWindowRect(hwnd, out r);
                int w = r.R - r.L, h = r.B - r.T;
                if (w <= 0 || h <= 0) return false;
                using (var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                {
                    bool ok;
                    using (var g = Graphics.FromImage(bmp))
                    {
                        IntPtr hdc = g.GetHdc();
                        ok = PrintWindow(hwnd, hdc, flags);
                        g.ReleaseHdc(hdc);
                    }
                    bmp.Save(pngPath, ImageFormat.Png);
                    return ok;
                }
            });
        }

        static IntPtr LP(int x, int y) { return new IntPtr((y << 16) | (x & 0xFFFF)); }
        // Mouse input posted straight to a window (client coordinates); the real cursor is not moved.
        public static bool PostClick(string desk, IntPtr hwnd, int x, int y)
        {
            return On(desk, delegate()
            {
                PostMessage(hwnd, 0x0200, IntPtr.Zero, LP(x, y));
                PostMessage(hwnd, 0x0201, new IntPtr(1), LP(x, y));
                return PostMessage(hwnd, 0x0202, IntPtr.Zero, LP(x, y));
            });
        }
        public static bool PostDrag(string desk, IntPtr hwnd, int x1, int y1, int x2, int y2)
        {
            return On(desk, delegate()
            {
                PostMessage(hwnd, 0x0200, IntPtr.Zero, LP(x1, y1));
                PostMessage(hwnd, 0x0201, new IntPtr(1), LP(x1, y1));
                for (int i = 1; i <= 8; i++) { PostMessage(hwnd, 0x0200, new IntPtr(1), LP(x1 + (x2 - x1) * i / 8, y1 + (y2 - y1) * i / 8)); Thread.Sleep(15); }
                return PostMessage(hwnd, 0x0202, IntPtr.Zero, LP(x2, y2));
            });
        }
        // press, keep the button down for holdMs, release
        public static bool PostHold(string desk, IntPtr hwnd, int x, int y, int holdMs)
        {
            return On(desk, delegate()
            {
                PostMessage(hwnd, 0x0200, IntPtr.Zero, LP(x, y));
                PostMessage(hwnd, 0x0201, new IntPtr(1), LP(x, y));
                Thread.Sleep(Math.Max(0, holdMs));
                return PostMessage(hwnd, 0x0202, IntPtr.Zero, LP(x, y));
            });
        }
        // press, wait holdMs, move to (x2, y2) in steps, release
        public static bool PostDragSlow(string desk, IntPtr hwnd, int x1, int y1, int x2, int y2, int steps, int stepMs, int holdMs)
        {
            return On(desk, delegate()
            {
                if (steps < 1) steps = 1;
                PostMessage(hwnd, 0x0200, IntPtr.Zero, LP(x1, y1));
                PostMessage(hwnd, 0x0201, new IntPtr(1), LP(x1, y1));
                Thread.Sleep(Math.Max(0, holdMs));
                for (int i = 1; i <= steps; i++) { PostMessage(hwnd, 0x0200, new IntPtr(1), LP(x1 + (x2 - x1) * i / steps, y1 + (y2 - y1) * i / steps)); Thread.Sleep(Math.Max(1, stepMs)); }
                return PostMessage(hwnd, 0x0202, IntPtr.Zero, LP(x2, y2));
            });
        }
        public static bool PostMove(string desk, IntPtr hwnd, int x, int y)
        {
            return On(desk, delegate() { return PostMessage(hwnd, 0x0200, IntPtr.Zero, LP(x, y)); });
        }
        // characters as the keyboard would send them (WM_CHAR)
        public static bool PostText(string desk, IntPtr hwnd, string text)
        {
            return On(desk, delegate()
            {
                bool ok = true;
                foreach (char ch in text ?? "") { ok = PostMessage(hwnd, 0x0102, new IntPtr(ch), new IntPtr(1)) && ok; Thread.Sleep(15); }
                return ok;
            });
        }
        public static bool PostKey(string desk, IntPtr hwnd, int vk)
        {
            return On(desk, delegate()
            {
                uint sc = MapVirtualKey((uint)vk, 0);
                PostMessage(hwnd, 0x0100, new IntPtr(vk), new IntPtr((long)(1 | (sc << 16))));
                return PostMessage(hwnd, 0x0101, new IntPtr(vk), new IntPtr((long)(1 | (sc << 16) | 0xC0000000)));
            });
        }

        [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object ppv);
        // Automation object behind an Office document window (OBJID_NATIVEOM), e.g. PowerPoint's DocumentWindow.
        public static object NativeOm(string desk, IntPtr hwnd)
        {
            return On(desk, delegate()
            {
                Guid iid = new Guid("00020400-0000-0000-C000-000000000046");
                object o; int hr = AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, ref iid, out o);
                if (hr != 0) throw new InvalidOperationException("AccessibleObjectFromWindow hr=0x" + hr.ToString("X8"));
                return o;
            });
        }

        public static IntPtr Foreground() { return GetForegroundWindow(); }
        public static bool SetForeground(IntPtr hwnd) { return SetForegroundWindow(hwnd); }
        public static void Show(IntPtr hwnd, int cmd) { ShowWindow(hwnd, cmd); }
        public static void MoveNoActivate(IntPtr hwnd, int x, int y) { SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0, 0x0001 | 0x0004 | 0x0010); }
    }

    // Running Object Table lookup that never launches a server (unlike GetObject / BindToMoniker).
    public static class Rot
    {
        [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
        [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        public static List<string> Names()
        {
            var list = new List<string>();
            IRunningObjectTable rot; GetRunningObjectTable(0, out rot);
            IEnumMoniker en; rot.EnumRunning(out en);
            IBindCtx ctx; CreateBindCtx(0, out ctx);
            var mk = new IMoniker[1];
            while (en.Next(1, mk, IntPtr.Zero) == 0)
            {
                string name; try { mk[0].GetDisplayName(ctx, null, out name); } catch { name = "?"; }
                list.Add(name);
            }
            return list;
        }

        public static object GetByName(string displayName)
        {
            IRunningObjectTable rot; GetRunningObjectTable(0, out rot);
            IEnumMoniker en; rot.EnumRunning(out en);
            IBindCtx ctx; CreateBindCtx(0, out ctx);
            var mk = new IMoniker[1];
            while (en.Next(1, mk, IntPtr.Zero) == 0)
            {
                string name; try { mk[0].GetDisplayName(ctx, null, out name); } catch { continue; }
                if (string.Equals(name, displayName, StringComparison.OrdinalIgnoreCase))
                {
                    object o; if (rot.GetObject(mk[0], out o) == 0) return o;
                }
            }
            return null;
        }
    }
}
