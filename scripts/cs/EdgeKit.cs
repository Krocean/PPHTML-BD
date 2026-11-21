// Drives Microsoft Edge (or Chrome) without a window, through its DevTools protocol on a loopback port.
// The skill uses the browser as a build tool only: converting modern CSS / JavaScript, subsetting fonts,
// drawing reference pictures and checking the web version. Nothing is installed and no page of the user's
// browser profile is touched (a throw-away profile folder is used).
// C# 5 only (compiled by Windows PowerShell's Add-Type). References: System, System.Web.Extensions.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.Win32;

namespace EdgeKit
{
    public class Session : IDisposable
    {
        Process proc;
        ClientWebSocket ws;
        int nextId;
        string profile;
        readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue, RecursionLimit = 400 };
        public List<string> Errors = new List<string>();   // uncaught exceptions and console errors of the page
        public string Exe = "";
        public string Browser = "";

        public static string FindBrowser()
        {
            var cands = new List<string>();
            foreach (var root in new[] { Environment.GetEnvironmentVariable("ProgramFiles(x86)"), Environment.GetEnvironmentVariable("ProgramFiles"), Environment.GetEnvironmentVariable("LOCALAPPDATA") })
            {
                if (string.IsNullOrEmpty(root)) continue;
                cands.Add(Path.Combine(root, @"Microsoft\Edge\Application\msedge.exe"));
            }
            foreach (var key in new[] { @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe", @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe" })
            {
                try { using (var k = Registry.LocalMachine.OpenSubKey(key)) { if (k != null) { var v = k.GetValue(null) as string; if (!string.IsNullOrEmpty(v)) cands.Add(v.Trim('"')); } } } catch { }
            }
            foreach (var root in new[] { Environment.GetEnvironmentVariable("ProgramFiles"), Environment.GetEnvironmentVariable("ProgramFiles(x86)"), Environment.GetEnvironmentVariable("LOCALAPPDATA") })
            {
                if (string.IsNullOrEmpty(root)) continue;
                cands.Add(Path.Combine(root, @"Google\Chrome\Application\chrome.exe"));
            }
            foreach (var c in cands) { try { if (File.Exists(c)) return c; } catch { } }
            return null;
        }

        public Session(string exe, string profileDir, int width, int height, int startTimeoutMs)
        {
            Exe = exe; profile = profileDir;
            Directory.CreateDirectory(profile);
            var portFile = Path.Combine(profile, "DevToolsActivePort");
            try { if (File.Exists(portFile)) File.Delete(portFile); } catch { }
            var args = "--headless=new --remote-debugging-port=0 --user-data-dir=\"" + profile + "\" --no-first-run --no-default-browser-check --disable-extensions --disable-sync"
                + " --disable-background-networking --disable-component-update --disable-default-apps --disable-features=Translate,OptimizationHints,MediaRouter"
                + " --hide-scrollbars --mute-audio --force-device-scale-factor=1 --window-size=" + width + "," + height + " about:blank";
            var psi = new ProcessStartInfo(exe, args) { UseShellExecute = false, CreateNoWindow = true };
            proc = Process.Start(psi);
            var sw = Stopwatch.StartNew();
            int port = 0;
            while (sw.ElapsedMilliseconds < startTimeoutMs)
            {
                Thread.Sleep(100);
                try
                {
                    if (File.Exists(portFile))
                    {
                        var lines = File.ReadAllLines(portFile);
                        if (lines.Length > 0 && int.TryParse(lines[0].Trim(), out port) && port > 0) break;
                    }
                }
                catch { }
                if (proc.HasExited) throw new Exception("the browser ended at once (exit code " + proc.ExitCode + ")");
            }
            if (port == 0) { Kill(); throw new Exception("the browser did not open its DevTools port in time"); }
            string wsUrl = null;
            while (sw.ElapsedMilliseconds < startTimeoutMs && wsUrl == null)
            {
                try
                {
                    var list = json.DeserializeObject(Http("http://127.0.0.1:" + port + "/json/list")) as object[];
                    if (list != null)
                    {
                        foreach (var o in list)
                        {
                            var d = o as Dictionary<string, object>;
                            if (d != null && (string)d["type"] == "page" && d.ContainsKey("webSocketDebuggerUrl")) { wsUrl = (string)d["webSocketDebuggerUrl"]; break; }
                        }
                    }
                }
                catch { }
                if (wsUrl == null) Thread.Sleep(150);
            }
            if (wsUrl == null) { Kill(); throw new Exception("the browser did not list a page to attach to"); }
            try { Browser = (string)((Dictionary<string, object>)json.DeserializeObject(Http("http://127.0.0.1:" + port + "/json/version")))["Browser"]; } catch { }
            ws = new ClientWebSocket();
            try { ws.Options.Proxy = null; } catch { }
            ws.Options.KeepAliveInterval = TimeSpan.FromSeconds(20);
            if (!ws.ConnectAsync(new Uri(wsUrl), CancellationToken.None).Wait(Math.Max(3000, startTimeoutMs - (int)sw.ElapsedMilliseconds))) { Kill(); throw new Exception("could not connect to the browser's DevTools socket"); }
            Send("Page.enable", null, 10000);
            Send("Runtime.enable", null, 10000);
            Send("Log.enable", null, 10000);
        }

        static string Http(string url)
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Proxy = null; req.Timeout = 3000; req.ReadWriteTimeout = 3000;
            using (var res = req.GetResponse())
            using (var sr = new StreamReader(res.GetResponseStream(), Encoding.UTF8)) return sr.ReadToEnd();
        }

        Dictionary<string, object> Receive(int timeoutMs)
        {
            var ms = new MemoryStream();
            var buf = new ArraySegment<byte>(new byte[1 << 16]);
            using (var cts = new CancellationTokenSource(Math.Max(1, timeoutMs)))
            {
                while (true)
                {
                    WebSocketReceiveResult r;
                    try { r = ws.ReceiveAsync(buf, cts.Token).GetAwaiter().GetResult(); }
                    catch (OperationCanceledException) { return null; }
                    if (r.MessageType == WebSocketMessageType.Close) throw new Exception("the browser closed the connection");
                    ms.Write(buf.Array, 0, r.Count);
                    if (r.EndOfMessage) break;
                }
            }
            return json.DeserializeObject(Encoding.UTF8.GetString(ms.GetBuffer(), 0, (int)ms.Length)) as Dictionary<string, object>;
        }

        void HandleEvent(Dictionary<string, object> m)
        {
            object mo; if (!m.TryGetValue("method", out mo)) return;
            var method = mo as string; var p = m.ContainsKey("params") ? m["params"] as Dictionary<string, object> : null;
            if (p == null) return;
            try
            {
                if (method == "Runtime.exceptionThrown")
                {
                    var d = p["exceptionDetails"] as Dictionary<string, object>;
                    Errors.Add(Describe(d));
                }
                else if (method == "Runtime.consoleAPICalled" && (string)p["type"] == "error")
                {
                    var sb = new StringBuilder("console.error:");
                    foreach (var a in (object[])p["args"])
                    {
                        var ad = a as Dictionary<string, object>; if (ad == null) continue;
                        object v; if (ad.TryGetValue("value", out v) && v != null) sb.Append(' ').Append(v); else if (ad.TryGetValue("description", out v) && v != null) sb.Append(' ').Append(v);
                    }
                    Errors.Add(sb.ToString());
                }
                else if (method == "Log.entryAdded")
                {
                    var e = p["entry"] as Dictionary<string, object>;
                    if (e != null && (string)e["level"] == "error") Errors.Add((string)e["source"] + ": " + (string)e["text"] + (e.ContainsKey("url") ? " (" + e["url"] + ")" : ""));
                }
            }
            catch { }
        }

        static string Describe(Dictionary<string, object> d)
        {
            if (d == null) return "script error";
            var text = d.ContainsKey("text") ? Convert.ToString(d["text"]) : "script error";
            object ex; if (d.TryGetValue("exception", out ex) && ex is Dictionary<string, object>)
            {
                object desc; if (((Dictionary<string, object>)ex).TryGetValue("description", out desc) && desc != null) text = Convert.ToString(desc);
            }
            var first = text.Split('\n')[0];
            object line; if (d.TryGetValue("lineNumber", out line)) first += " (line " + (Convert.ToInt32(line) + 1) + ")";
            return first;
        }

        public Dictionary<string, object> Send(string method, object prm, int timeoutMs)
        {
            int id = ++nextId;
            var msg = new Dictionary<string, object>(); msg["id"] = id; msg["method"] = method; if (prm != null) msg["params"] = prm;
            var bytes = Encoding.UTF8.GetBytes(json.Serialize(msg));
            ws.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None).GetAwaiter().GetResult();
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                var m = Receive(timeoutMs - (int)sw.ElapsedMilliseconds);
                if (m == null) break;
                object mid;
                if (m.TryGetValue("id", out mid) && Convert.ToInt32(mid) == id)
                {
                    if (m.ContainsKey("error")) throw new Exception(method + ": " + json.Serialize(m["error"]));
                    var res = m.ContainsKey("result") ? m["result"] as Dictionary<string, object> : null;
                    return res ?? new Dictionary<string, object>();
                }
                HandleEvent(m);
            }
            throw new TimeoutException(method + " did not answer within " + timeoutMs + " ms");
        }

        // Evaluates an expression in the page; the value comes back as text (objects as JSON). Throws on a script error.
        public string Eval(string expression, bool awaitPromise, int timeoutMs)
        {
            var prm = new Dictionary<string, object>();
            prm["expression"] = expression; prm["returnByValue"] = true; prm["awaitPromise"] = awaitPromise; prm["userGesture"] = true;
            var r = Send("Runtime.evaluate", prm, timeoutMs);
            if (r.ContainsKey("exceptionDetails")) throw new Exception(Describe(r["exceptionDetails"] as Dictionary<string, object>));
            var res = r.ContainsKey("result") ? r["result"] as Dictionary<string, object> : null;
            object v;
            if (res != null && res.TryGetValue("value", out v) && v != null) return v as string ?? json.Serialize(v);
            return null;
        }

        // Runs a script file of the skill in the page (the text is sent over the socket: the browser needs no file access).
        public void LoadScript(string path, int timeoutMs)
        {
            Eval(File.ReadAllText(path, Encoding.UTF8) + "\n;void 0", false, timeoutMs);
        }

        // Sets a global variable of the page to a string (any size; sent in pieces).
        public void SetText(string name, string text, int timeoutMs)
        {
            const int piece = 4 * 1024 * 1024;
            Eval("window[" + json.Serialize(name) + "] = []; void 0", false, timeoutMs);
            for (int i = 0; i < text.Length; i += piece)
            {
                var part = text.Substring(i, Math.Min(piece, text.Length - i));
                Eval("window[" + json.Serialize(name) + "].push(" + json.Serialize(part) + "); void 0", false, timeoutMs);
            }
            Eval("window[" + json.Serialize(name) + "] = window[" + json.Serialize(name) + "].join(''); void 0", false, timeoutMs);
        }

        public void SetViewport(int width, int height, double scale)
        {
            var prm = new Dictionary<string, object>();
            prm["width"] = width; prm["height"] = height; prm["deviceScaleFactor"] = scale; prm["mobile"] = false;
            Send("Emulation.setDeviceMetricsOverride", prm, 10000);
        }

        // Replaces the page with the given document (scripts run) and waits until it has loaded.
        public bool SetContent(string html, int timeoutMs)
        {
            var nav = new Dictionary<string, object>(); nav["url"] = "about:blank";
            Send("Page.navigate", nav, 10000);
            WaitReady(5000);
            var tree = Send("Page.getFrameTree", null, 10000);
            var frame = (Dictionary<string, object>)((Dictionary<string, object>)tree["frameTree"])["frame"];
            var prm = new Dictionary<string, object>(); prm["frameId"] = frame["id"]; prm["html"] = html;
            Send("Page.setDocumentContent", prm, Math.Max(10000, timeoutMs));
            return WaitReady(timeoutMs);
        }

        public bool Navigate(string url, int timeoutMs)
        {
            var nav = new Dictionary<string, object>(); nav["url"] = url;
            Send("Page.navigate", nav, Math.Max(10000, timeoutMs));
            return WaitReady(timeoutMs);
        }

        public bool WaitReady(int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try { if (Eval("document.readyState", false, 5000) == "complete") return true; } catch (TimeoutException) { throw; } catch { }
                Thread.Sleep(60);
            }
            return false;
        }

        public void Screenshot(string path)
        {
            var prm = new Dictionary<string, object>(); prm["format"] = "png"; prm["captureBeyondViewport"] = false;
            var r = Send("Page.captureScreenshot", prm, 30000);
            File.WriteAllBytes(path, Convert.FromBase64String((string)r["data"]));
        }

        void Kill()
        {
            try { if (proc != null && !proc.HasExited) proc.Kill(); } catch { }
        }

        public void Dispose()
        {
            try { if (ws != null && ws.State == WebSocketState.Open) { var id = ++nextId; var bytes = Encoding.UTF8.GetBytes("{\"id\":" + id + ",\"method\":\"Browser.close\"}"); ws.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None).Wait(2000); } } catch { }
            try { if (proc != null) proc.WaitForExit(4000); } catch { }
            Kill();
            try { if (ws != null) ws.Dispose(); } catch { }
            for (int i = 0; i < 6; i++)
            {
                try { if (Directory.Exists(profile)) Directory.Delete(profile, true); break; } catch { Thread.Sleep(400); }
            }
        }
    }
}
