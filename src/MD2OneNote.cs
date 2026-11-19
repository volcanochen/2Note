// MD2OneNote.Addin - OneNote COM add-in (per-user, no admin needed)
// Adds "Paste as OneNote" to page context menu + Home tab, and a Ctrl+Alt+M hotkey.
// Reads clipboard 'HTML Format' as raw bytes (UTF-8) and converts to NATIVE OneNote
// XML (one:OE / one:OEChildren), bypassing OneNote's HTML importer so nesting survives.
// Conversion logic is a 1:1 port of md2onexml.ps1 v4 (E2E-tested).

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

[assembly: System.Reflection.AssemblyVersion("1.1.0.0")]

namespace MD2OneNote
{
    // ---------- Office / VS interop (OneMore-style NO-PIA type-equivalent declarations) ----------

    // ---------- Office / VS interop (OneMore-style NO-PIA type-equivalent declarations) ----------
    // 2026-11-05: plain [ComImport] declarations crashed at OnConnection dispatch (clr.dll
    // 0x57ce3b AV / CLR fail-fast 80131506) while OneMore survived the identical call in the
    // identical DllSurrogate. OneMore declares these types as no-PIA embedded equivalents
    // ([TypeIdentifier] + [CompilerGenerated] + [Guid]) so the CLR builds CCW stubs from
    // metadata and never consults the (broken-on-this-machine) typelib registry. Exact GUIDs
    // of enums came from OneMore's assembly which mirrors the real Add-In Designer typelib.

    [CompilerGenerated, Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")]
    [TypeIdentifier("ac0714f2-3d04-11d1-ae7d-00a0c90f26f4", "Extensibility.ext_ConnectMode")]
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    [CompilerGenerated, Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")]
    [TypeIdentifier("ac0714f2-3d04-11d1-ae7d-00a0c90f26f4", "Extensibility.ext_DisconnectMode")]
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    // IID B65AD801 = Extensibility._IDTExtensibility2 (vtable order verified against
    // ksaddndr.dll typelib slots 56/64/72/80/88 and OneMore's working declaration).
    // 2026-11-05 ROOT CAUSE FIX: parameter marshaling attrs are MANDATORY. Without
    // MarshalAs(IDispatch) an `object` param defaults to VARIANT (24 bytes / 3 stack slots
    // on x64) while the real interface passes a raw IDispatch* (8 bytes / 1 slot) - the
    // stub then consumes the wrong slots -> AV (clr.dll 0x57ce3b) or
    // InvalidOleVariantTypeException (0x80131531) -> OneNote "runtime error". OneMore has
    // exactly these attrs (verified via reflection 2026-11-05).
    [ComImport, CompilerGenerated, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), TypeIdentifier]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection(
            [In, MarshalAs(UnmanagedType.IDispatch)] object Application,
            [In] ext_ConnectMode ConnectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(
            [In] ext_DisconnectMode RemoveMode,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    // IID 000C0396 = Microsoft.Office.Core.IRibbonExtensibility
    [ComImport, CompilerGenerated, Guid("000C0396-0000-0000-C000-000000000046"), TypeIdentifier]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    // ---------- main add-in class ----------

    [ComVisible(true)]
    [Guid("4A9D4FBD-D100-41EF-9C9C-7CD1DB2EE81B")]
    [ProgId("MD2OneNote.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        // 2026-11-05: the host's Application object is a pure dispinterface (real IID
        // {452AC71A-B655-4967-A208-A4CC39DD7949}); QI for the PIA IApplication IID
        // {A47D3223-...} fails E_NOINTERFACE (proven in-proc via addin.log 15:59/16:16).
        // Keep the raw RCW; all calls go through the hand-rolled IDispatch invoker below.
        private static object _appObj;
        private static HotkeyWindow _hotkey;

        static Connect()
        {
            Log("static ctor: assembly loaded in-process");
        }

        public Connect()
        {
            Log("instance ctor");
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _appObj = Application;
            Log("OnConnection mode=" + ConnectMode + " appObj=" +
                (Application != null ? Application.GetType().FullName : "null"));
            EnsureHotkey();
        }

        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete");
            EnsureHotkey();
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("OnDisconnection mode=" + RemoveMode);
            if (_hotkey != null) { _hotkey.Dispose(); _hotkey = null; }
            // 2026-11-06 zombie hygiene: drop our COM reference so OneNote can exit
            ReleaseAppObject("OnDisconnection");
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom)
        {
            if (_hotkey != null) { _hotkey.Dispose(); _hotkey = null; }
            ReleaseAppObject("OnBeginShutdown");
        }

        private static void ReleaseAppObject(string where)
        {
            if (_appObj == null) return;
            try { Marshal.FinalReleaseComObject(_appObj); }
            catch (Exception ex) { Log("release appObj error (" + where + "): " + ex.Message); }
            _appObj = null;
            Log("appObj released (" + where + ")");
        }

        private static void EnsureHotkey()
        {
            if (_hotkey != null) return;
            // 2026-11-05: configurable via C:\ProgramData\MD2OneNote\hotkey.txt
            // (see HotkeyConfig; set with md2addin\set-hotkey.ps1).
            string spec = HotkeyConfig.ReadSpec();
            uint mod, vk; string norm;
            if (!HotkeyConfig.TryParse(spec, out mod, out vk, out norm) || norm == "")
            {
                mod = 3u; vk = 0x4Du; norm = "Ctrl+Alt+M";   // MOD_CONTROL|MOD_ALT, 'M'
                if (spec != "") Log("hotkey config '" + spec + "' invalid -> default Ctrl+Alt+M");
            }
            try
            {
                // NOTE: fires on the dedicated hotkey STA thread -> cannot use _appObj
                // (RCW created on the host's STA). Run(null) resolves via GetActiveObject.
                // allowCreate=false: never spawn a headless OneNote instance from the
                // hotkey path (zombie-process source, fixed 2026-11-06).
                _hotkey = new HotkeyWindow(mod, vk, delegate
                {
                    try { PasteEngine.Run(null, null, false); }
                    catch (Exception ex) { Log("hotkey run error: " + ex.Message); }
                });
                Log("hotkey " + norm + " registered (config: " + (spec == "" ? "<none>" : spec) + ")");
            }
            catch (Exception ex)
            {
                Log("hotkey register FAILED for " + norm + ": " + ex.Message +
                    " (owned by another app; menu button still works - change hotkey.txt)");
            }
        }

        // ---------- ribbon ----------

        [ComVisible(true)]
        public string GetCustomUI(string RibbonID)
        {
            Log("GetCustomUI called, RibbonID=" + (RibbonID ?? ""));
            // 2026-11-05: structure mirrors OneMore's PROVEN-WORKING customUI (saved at
            // C:\ProgramData\MD2OneNote\onemore-customui.xml): root onLoad callback +
            // <ribbon> with a REAL tab (empty <tabs/> is the prime suspect for the
            // context menu being rejected wholesale) + <contextMenus> anchored at
            // insertBeforeMso="RenamePageOneNote" (OneMore-proven anchor).
            return
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnRibbonLoad\">" +
                "<ribbon><tabs>" +
                "<tab id=\"Tab2Note\" insertAfterMso=\"TabView\" label=\"2note\">" +
                "<group id=\"grp2NotePaste\" label=\"Paste\">" +
                "<button id=\"btn2NotePaste\" label=\"Paste as better structure\" imageMso=\"Paste\" " +
                "onAction=\"OnPaste\" size=\"large\"/>" +
                "</group>" +
                "<group id=\"grp2NoteAbout\" label=\"About\">" +
                "<button id=\"btn2NoteAbout\" label=\"\u5173\u4E8E 2Note\" " +
                "onAction=\"OnAbout\" size=\"large\"/>" +
                "</group></tab>" +
                "</tabs></ribbon>" +
                "<contextMenus><contextMenu idMso=\"ContextMenuPageArea\">" +
                "<button id=\"ctx2Note\" label=\"\u7C98\u8D34\u4E3A OneNote \u7ED3\u6784 (Alt+Shift+V)\" " +
                "imageMso=\"Paste\" onAction=\"OnPaste\" insertBeforeMso=\"RenamePageOneNote\"/>" +
                "</contextMenu></contextMenus></customUI>";
        }

        // ribbon onAction callback (invoked via IDispatch)
        public void OnPaste(object control)
        {
            try { PasteEngine.Run(_appObj, null); }
            catch (Exception ex) { Log("OnPaste error: " + ex.Message); }
        }

        // ---- About dialog (imitates OneMore's About box; 2026-11-06) ----

        public void OnAbout(object control)
        {
            try { ShowAboutDialog(); }
            catch (Exception ex) { Log("OnAbout error: " + ex.Message); }
        }

        private void ShowAboutDialog()
        {
            string ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(3);
            string oneVer = "";
            try
            {
                OneNoteApp app = OneNoteApp.Wrap(_appObj);
                if (app != null) oneVer = app.HostVersion() ?? "";
            }
            catch (Exception ex) { Log("host version failed: " + ex.Message); }

            System.Windows.Forms.Form f = new System.Windows.Forms.Form();
            f.Text = "\u5173\u4E8E 2Note \u52A0\u8F7D\u9879";
            f.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            f.MaximizeBox = false; f.MinimizeBox = false;
            f.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            f.ClientSize = new System.Drawing.Size(640, 400);
            f.BackColor = System.Drawing.Color.White;
            f.Font = new System.Drawing.Font("Microsoft YaHei UI", 9.5f);

            System.Windows.Forms.Label title = new System.Windows.Forms.Label();
            title.Text = "\u9002\u7528 OneNote \u7684 2Note \u52A0\u8F7D\u9879";
            title.Font = new System.Drawing.Font("Microsoft YaHei UI", 15f, System.Drawing.FontStyle.Bold);
            title.Location = new System.Drawing.Point(24, 28);
            title.AutoSize = true;
            f.Controls.Add(title);

            System.Windows.Forms.Label verLine = new System.Windows.Forms.Label();
            verLine.Text = "\u63D2\u4EF6\u7248\u672C " + ver + "  \u5B89\u88C5\u5728 OneNote\u7248\u672C " + (oneVer == "" ? "?" : oneVer);
            verLine.Location = new System.Drawing.Point(26, 74);
            verLine.AutoSize = true;
            verLine.ForeColor = System.Drawing.Color.FromArgb(60, 60, 60);
            f.Controls.Add(verLine);

            // purple "2Note" badge - OneNote brand purple (#7719AA), anti-aliased,
            // filled squircle with vertical gradient (2026-11-06 redesign)
            System.Windows.Forms.Panel badge = new System.Windows.Forms.Panel();
            badge.Size = new System.Drawing.Size(130, 130);
            badge.Location = new System.Drawing.Point(640 - 148, 12);
            badge.Paint += delegate(object s, System.Windows.Forms.PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                System.Drawing.Rectangle rect = new System.Drawing.Rectangle(1, 1, badge.Width - 3, badge.Height - 3);
                using (System.Drawing.Drawing2D.GraphicsPath gp = RoundedRect(rect, 30))
                using (System.Drawing.Drawing2D.LinearGradientBrush br = new System.Drawing.Drawing2D.LinearGradientBrush(
                    rect, System.Drawing.Color.FromArgb(154, 63, 208), System.Drawing.Color.FromArgb(119, 25, 170), 90f))
                {
                    e.Graphics.FillPath(br, gp);
                }
                string txt = "2Note";
                using (System.Drawing.Font fnt = FitFont(e.Graphics, txt, badge.Width - 20, 30f))
                {
                    System.Drawing.SizeF ts = e.Graphics.MeasureString(txt, fnt);
                    e.Graphics.DrawString(txt, fnt, System.Drawing.Brushes.White,
                        (badge.Width - ts.Width) / 2f, (badge.Height - ts.Height) / 2f - 2f);
                }
            };
            f.Controls.Add(badge);

            System.Windows.Forms.Panel body = new System.Windows.Forms.Panel();
            body.BackColor = System.Drawing.Color.FromArgb(243, 243, 243);
            body.Location = new System.Drawing.Point(0, 148);
            body.Size = new System.Drawing.Size(640, 252);
            f.Controls.Add(body);

            int y = 20;
            System.Windows.Forms.Label copyright = new System.Windows.Forms.Label();
            copyright.Text = "\u7248\u6743\u6240\u6709 @ 2026 volcanochen\u3002";
            copyright.Location = new System.Drawing.Point(24, y);
            copyright.AutoSize = true;
            copyright.ForeColor = System.Drawing.Color.FromArgb(80, 80, 80);
            body.Controls.Add(copyright); y += 34;

            System.Windows.Forms.LinkLabel check = new System.Windows.Forms.LinkLabel();
            check.Text = "\u68C0\u67E5\u66F4\u65B0";
            check.Location = new System.Drawing.Point(24, y);
            check.AutoSize = true;
            check.LinkArea = new System.Windows.Forms.LinkArea(0, check.Text.Length);
            check.LinkClicked += delegate { CheckForUpdate(f); };
            body.Controls.Add(check); y += 36;

            System.Windows.Forms.LinkLabel gh = new System.Windows.Forms.LinkLabel();
            gh.Text = "GitHub \u9879\u76EE: https://github.com/volcanochen/2Note";
            gh.Location = new System.Drawing.Point(24, y);
            gh.AutoSize = true;
            gh.LinkArea = new System.Windows.Forms.LinkArea(0, gh.Text.Length);
            gh.LinkClicked += delegate
            {
                try { System.Diagnostics.Process.Start("https://github.com/volcanochen/2Note"); }
                catch { }
            };
            body.Controls.Add(gh); y += 36;

            string logPath = PasteEngine.LogPath;
            System.Windows.Forms.LinkLabel logLink = new System.Windows.Forms.LinkLabel();
            logLink.Text = logPath;
            logLink.Location = new System.Drawing.Point(24, y);
            logLink.AutoSize = true;
            logLink.LinkArea = new System.Windows.Forms.LinkArea(0, logLink.Text.Length);
            logLink.LinkClicked += delegate
            {
                try { System.Diagnostics.Process.Start("notepad.exe", logPath); } catch { }
            };
            body.Controls.Add(logLink); y += 36;

            System.Windows.Forms.LinkLabel clear = new System.Windows.Forms.LinkLabel();
            clear.Text = "\u6E05\u9664\u65E5\u5FD7\u6587\u4EF6";
            clear.Location = new System.Drawing.Point(24, y);
            clear.AutoSize = true;
            clear.LinkArea = new System.Windows.Forms.LinkArea(0, clear.Text.Length);
            clear.LinkClicked += delegate
            {
                try
                {
                    System.IO.File.Delete(PasteEngine.LogPath);
                    System.Windows.Forms.MessageBox.Show(f, "\u65E5\u5FD7\u5DF2\u6E05\u9664\u3002");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(f, "\u6E05\u9664\u5931\u8D25\uFF1A" + ex.Message);
                }
            };
            body.Controls.Add(clear);

            System.Windows.Forms.Button ok = new System.Windows.Forms.Button();
            ok.Text = "\u786E\u5B9A";
            ok.Size = new System.Drawing.Size(110, 32);
            ok.Location = new System.Drawing.Point(640 - 132, 400 - 46);
            ok.DialogResult = System.Windows.Forms.DialogResult.OK;
            f.Controls.Add(ok);
            f.AcceptButton = ok;

            f.ShowDialog();
            f.Dispose();
        }

        // ---- update check against GitHub Releases (2026-11-06) ----

        private void CheckForUpdate(System.Windows.Forms.Form owner)
        {
            // runs on a worker thread; results marshalled back to the UI thread
            System.Windows.Forms.MethodInvoker done = delegate
            {
                // replaced below via closure locals - see worker
            };
            System.Threading.ThreadPool.QueueUserWorkItem(delegate
            {
                string ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(3);
                string latest = "", url = "", notes = "", err = "";
                try
                {
                    System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(
                        "https://api.github.com/repos/volcanochen/2Note/releases/latest");
                    req.UserAgent = "2Note-Addon";
                    req.Timeout = 10000;
                    using (System.Net.HttpWebResponse resp = (System.Net.HttpWebResponse)req.GetResponse())
                    using (System.IO.StreamReader sr = new System.IO.StreamReader(resp.GetResponseStream()))
                    {
                        string json = sr.ReadToEnd();
                        Match tm = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"([^\"]+)\"");
                        if (tm.Success) latest = tm.Groups[1].Value;
                        Match um = Regex.Match(json, "\"html_url\"\\s*:\\s*\"([^\"]+)\"");
                        if (um.Success) url = um.Groups[1].Value;
                        Match nm = Regex.Match(json, "\"body\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
                        if (nm.Success)
                        {
                            notes = nm.Groups[1].Value
                                .Replace("\\r\\n", "\n").Replace("\\n", "\n").Replace("\\r", "")
                                .Replace("\\\"", "\"").Replace("\\\\", "\\");
                            if (notes.Length > 600) notes = notes.Substring(0, 600) + "...";
                        }
                    }
                }
                catch (Exception ex) { err = ex.Message; }

                owner.BeginInvoke((System.Windows.Forms.MethodInvoker)delegate
                {
                    if (err.Length > 0)
                    {
                        System.Windows.Forms.MessageBox.Show(owner,
                            "\u68C0\u67E5\u66F4\u65B0\u5931\u8D25\uFF1A" + err);
                        return;
                    }
                    string cur = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(3);
                    int cmp = CompareVersions(cur, latest);
                    if (cmp < 0)
                    {
                        string msg = "\u53D1\u73B0\u65B0\u7248\u672C " + latest + "\uFF08\u5F53\u524D " + cur + "\uFF09";
                        if (notes.Length > 0) msg += "\n\n" + notes;
                        msg += "\n\n\u662F\u5426\u6253\u5F00\u4E0B\u8F7D\u9875\u9762\uFF1F";
                        if (System.Windows.Forms.MessageBox.Show(owner, msg,
                            "2Note \u66F4\u65B0", System.Windows.Forms.MessageBoxButtons.YesNo,
                            System.Windows.Forms.MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.Yes)
                        {
                            try { System.Diagnostics.Process.Start(url.Length > 0 ? url : "https://github.com/volcanochen/2Note/releases"); } catch { }
                        }
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show(owner,
                            "\u5DF2\u662F\u6700\u65B0\u7248\u672C\uFF08" + cur + "\uFF09\u3002");
                    }
                });
            });
        }

        // >0 when remote (b) is newer than local (a); tags like "v1.1.0" supported
        private static int CompareVersions(string a, string b)
        {
            Version va, vb;
            if (!Version.TryParse(a.TrimStart('v', 'V'), out va)) return 0;
            if (!Version.TryParse(b.TrimStart('v', 'V'), out vb)) return 0;
            return vb.CompareTo(va);
        }

        private static System.Drawing.Drawing2D.GraphicsPath RoundedRect(System.Drawing.Rectangle r, int rad)
        {
            System.Drawing.Drawing2D.GraphicsPath gp = new System.Drawing.Drawing2D.GraphicsPath();
            int d = rad * 2;
            gp.AddArc(r.X, r.Y, d, d, 180, 90);
            gp.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            gp.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            gp.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            gp.CloseFigure();
            return gp;
        }

        private static System.Drawing.Font FitFont(System.Drawing.Graphics g, string s, int maxW, float startSize)
        {
            float size = startSize;
            while (size > 10f)
            {
                System.Drawing.Font test = new System.Drawing.Font("Segoe UI", size, System.Drawing.FontStyle.Bold);
                float w = g.MeasureString(s, test).Width;
                test.Dispose();
                if (w <= maxW) return new System.Drawing.Font("Segoe UI", size, System.Drawing.FontStyle.Bold);
                size -= 2f;
            }
            return new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold);
        }

        // customUI onLoad callback (OneMore-style root callback; presence proven to work)
        public void OnRibbonLoad(object ribbonUI)
        {
            Log("OnRibbonLoad ribbonUI=" + (ribbonUI != null ? ribbonUI.GetType().FullName : "null"));
        }

        internal static void Log(string msg)
        {
            try
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MD2OneNote");
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, "addin.log"),
                    "[" + DateTime.Now.ToString("HH:mm:ss.fff") + "] " + msg + "\r\n", Encoding.UTF8);
            }
            catch { }
        }
    }

    // ---------- hotkey configuration (config file + parser) ----------

    // Config file: <dll dir>\hotkey.txt. First non-empty, non-# line is the combo,
    // e.g. "Ctrl+Alt+P", "Ctrl+Shift+V", "Win+F9". Missing/invalid file -> Ctrl+Alt+M.
    internal static class HotkeyConfig
    {
        public static string ConfigPath()
        {
            try
            {
                string dir = Path.GetDirectoryName(typeof(Connect).Assembly.Location);
                return Path.Combine(dir ?? "", "hotkey.txt");
            }
            catch { return "hotkey.txt"; }
        }

        public static string ReadSpec()
        {
            try
            {
                string p = ConfigPath();
                if (!File.Exists(p)) return "";
                foreach (string line in File.ReadAllLines(p, Encoding.UTF8))
                {
                    string t = (line ?? "").Trim();
                    if (t.Length == 0 || t.StartsWith("#")) continue;
                    return t;
                }
            }
            catch (Exception ex) { PasteEngine.Log("hotkey config read error: " + ex.Message); }
            return "";
        }

        // parse "Ctrl+Alt+P" style spec -> (mod, vk). Returns false on garbage.
        public static bool TryParse(string spec, out uint mod, out uint vk, out string norm)
        {
            mod = 0; vk = 0; norm = "";
            if (string.IsNullOrWhiteSpace(spec)) return false;
            uint m = 0; string key = null;
            foreach (string raw in spec.Split('+'))
            {
                string u = raw.Trim().ToUpperInvariant();
                if (u.Length == 0) return false;
                if (u == "CTRL" || u == "CONTROL") m |= 2u;
                else if (u == "ALT") m |= 1u;
                else if (u == "SHIFT") m |= 4u;
                else if (u == "WIN" || u == "WINDOWS") m |= 8u;
                else if (key == null) key = u;
                else return false;   // two non-modifier keys
            }
            if (key == null || m == 0) return false;
            uint v;
            if (key.Length == 1 && key[0] >= 'A' && key[0] <= 'Z') v = key[0];          // 0x41-0x5A
            else if (key.Length == 1 && key[0] >= '0' && key[0] <= '9') v = key[0];      // 0x30-0x39
            else if (key.Length >= 2 && key[0] == 'F')
            {
                uint fn;
                if (uint.TryParse(key.Substring(1), out fn) && fn >= 1 && fn <= 24)
                    v = 0x6Fu + fn; // F1=0x70
                else return false;
            }
            else
            {
                switch (key)
                {
                    case "SPACE": v = 0x20; break;
                    case "TAB": v = 0x09; break;
                    case "ENTER": v = 0x0D; break;
                    case "BACKSPACE": v = 0x08; break;
                    case "ESC": case "ESCAPE": v = 0x1B; break;
                    case "INSERT": v = 0x2D; break;
                    case "DELETE": case "DEL": v = 0x2E; break;
                    case "HOME": v = 0x24; break;
                    case "END": v = 0x23; break;
                    case "PAGEUP": v = 0x21; break;
                    case "PAGEDOWN": v = 0x22; break;
                    case "LEFT": v = 0x25; break;
                    case "UP": v = 0x26; break;
                    case "RIGHT": v = 0x27; break;
                    case "DOWN": v = 0x28; break;
                    default: return false;
                }
            }
            mod = m; vk = v;
            // normalized form for logging
            List<string> bits = new List<string>();
            if ((m & 2u) != 0) bits.Add("Ctrl");
            if ((m & 1u) != 0) bits.Add("Alt");
            if ((m & 4u) != 0) bits.Add("Shift");
            if ((m & 8u) != 0) bits.Add("Win");
            bits.Add(key.Length == 2 && key[0] == 'F' ? key.ToUpperInvariant() : key);
            norm = string.Join("+", bits);
            return true;
        }
    }

    // ---------- hotkey window (message-only, no WinForms dependency) ----------

    internal class HotkeyWindow : IDisposable
    {
        private const string ClassName = "MD2OneNoteHotkeyWnd";
        private const uint WM_HOTKEY = 0x0312;
        private delegate IntPtr WndProcDelegate(IntPtr hwnd, uint msg, IntPtr wp, IntPtr lp);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WNDCLASS
        {
            public uint style;
            public IntPtr lpfnWndProc;
            public int cbClsExtra;
            public int cbWndExtra;
            public IntPtr hInstance;
            public IntPtr hIcon;
            public IntPtr hCursor;
            public IntPtr hbrBackground;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpszMenuName;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpszClassName;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern ushort RegisterClassW(ref WNDCLASS wc);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateWindowExW(uint exStyle, string cls, string name, uint style,
            int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr inst, IntPtr param);
        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr h);
        [DllImport("user32.dll")] private static extern IntPtr DefWindowProcW(IntPtr h, uint m, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] private static extern bool RegisterHotKey(IntPtr h, int id, uint mod, uint vk);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr h, int id);
        [DllImport("user32.dll")] private static extern int GetMessage(out MSG m, IntPtr h, uint mn, uint mx);
        [DllImport("user32.dll")] private static extern bool TranslateMessage(ref MSG m);
        [DllImport("user32.dll")] private static extern IntPtr DispatchMessage(ref MSG m);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandleW(string name);
        [DllImport("kernel32.dll")] private static extern int GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool PostThreadMessage(int tid, uint msg, IntPtr wp, IntPtr lp);

        private const uint WM_QUIT = 0x0012;

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam;
            public uint time; public int ptX; public int ptY;
        }

        private static WndProcDelegate _keptProc;   // prevent GC of the callback
        private static bool _classRegistered;
        private IntPtr _hwnd;
        private readonly int _id = 1;
        private readonly Action _fired;
        private readonly uint _mod;
        private readonly uint _vk;
        private readonly System.Threading.AutoResetEvent _ready = new System.Threading.AutoResetEvent(false);
        private bool _ok;
        private int _err;
        private volatile int _tid;   // pump thread id, set before _ready is signalled

        // 2026-11-06 FIX: the window must live on a DEDICATED thread that pumps messages.
        // The previous version created the window on the host's STA thread with NO message
        // loop -> WM_HOTKEY was posted but never dispatched -> hotkey never fired (this was
        // the real reason "Ctrl+Alt+M didn't work", not an ownership conflict).
        public HotkeyWindow(uint modifiers, uint vk, Action fired)
        {
            _fired = fired;
            _mod = modifiers;
            _vk = vk;
            System.Threading.Thread t = new System.Threading.Thread(Pump);
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.IsBackground = true;
            t.Start();
            if (!_ready.WaitOne(4000))
                throw new TimeoutException("hotkey pump thread did not initialize in 4s");
            if (!_ok)
                throw new Win32Exception(_err, "RegisterHotKey failed (hotkey owned elsewhere)");
        }

        private void Pump()
        {
            _tid = GetCurrentThreadId();
            try
            {
                if (!_classRegistered)
                {
                    _keptProc = WndProc;
                    WNDCLASS wc = new WNDCLASS();
                    wc.lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_keptProc);
                    wc.hInstance = GetModuleHandleW(null);
                    wc.lpszClassName = ClassName;
                    if (RegisterClassW(ref wc) == 0)
                        throw new Win32Exception(Marshal.GetLastWin32Error(), "RegisterClassW failed");
                    _classRegistered = true;
                }
                _hwnd = CreateWindowExW(0, ClassName, ClassName, 0, 0, 0, 0, 0,
                    (IntPtr)(-3) /* HWND_MESSAGE */, IntPtr.Zero, GetModuleHandleW(null), IntPtr.Zero);
                if (_hwnd == IntPtr.Zero)
                {
                    _err = Marshal.GetLastWin32Error();
                    _ready.Set();
                    return;
                }
                _ok = RegisterHotKey(_hwnd, _id, _mod, _vk);
                if (!_ok) _err = Marshal.GetLastWin32Error();
                _ready.Set();
                if (!_ok) return;
                PasteEngine.Log("hotkey pump running on dedicated thread");
                MSG m;
                while (GetMessage(out m, IntPtr.Zero, 0, 0) > 0)
                {
                    TranslateMessage(ref m);
                    DispatchMessage(ref m);
                }
                // WM_QUIT received: unregister + destroy HERE on the owning thread
                // (UnregisterHotKey/DestroyWindow must run on the registering thread).
                if (_hwnd != IntPtr.Zero)
                {
                    UnregisterHotKey(_hwnd, _id);
                    DestroyWindow(_hwnd);
                    _hwnd = IntPtr.Zero;
                }
                PasteEngine.Log("hotkey pump exited");
            }
            catch (Exception ex)
            {
                PasteEngine.Log("hotkey pump error: " + ex.Message);
                _ready.Set();
            }
        }

        private IntPtr WndProc(IntPtr hwnd, uint msg, IntPtr wp, IntPtr lp)
        {
            if (msg == WM_HOTKEY)
            {
                PasteEngine.Log("WM_HOTKEY received (vk=" + ((uint)lp >> 16 & 0xFF).ToString() + ")");
                try { _fired(); }
                catch (Exception ex) { PasteEngine.Log("hotkey fire error: " + ex.Message); }
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        }

        public void Dispose()
        {
            // 2026-11-06 fix: old version only unregistered the hotkey but left the pump
            // thread running forever. Post WM_QUIT so GetMessage returns 0 and the pump
            // thread cleans up (unregister + destroy) and exits.
            if (_tid != 0)
            {
                PostThreadMessage(_tid, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
                PasteEngine.Log("hotkey pump stop requested (WM_QUIT -> tid " + _tid + ")");
                _tid = 0;
            }
        }
    }

    // ---------- raw clipboard access (no WinForms) ----------

    internal static class RawClipboard
    {
        [DllImport("user32.dll")] public static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")] public static extern bool CloseClipboard();
        [DllImport("user32.dll")] public static extern uint EnumClipboardFormats(uint iFormat);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClipboardFormatNameW(uint format, [Out] System.Text.StringBuilder name, int cch);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern uint RegisterClipboardFormatW(string lpszFormat);
        [DllImport("user32.dll")] public static extern IntPtr GetClipboardData(uint uFormat);
        [DllImport("kernel32.dll")] public static extern IntPtr GlobalLock(IntPtr hMem);
        [DllImport("kernel32.dll")] public static extern bool GlobalUnlock(IntPtr hMem);
        [DllImport("kernel32.dll")] public static extern UIntPtr GlobalSize(IntPtr hMem);

        public static byte[] GetFormatData(string formatName)
        {
            uint fmt = RegisterClipboardFormatW(formatName);
            PasteEngine.Log("clip '" + formatName + "': regId=" + fmt);
            for (int attempt = 0; attempt < 5; attempt++)
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    try
                    {
                        IntPtr h = GetClipboardData(fmt);
                        if (h == IntPtr.Zero)
                        {
                            PasteEngine.Log("clip '" + formatName + "': GetClipboardData=NULL (attempt " + attempt + ")");
                            return null;
                        }
                        IntPtr ptr = GlobalLock(h);
                        if (ptr == IntPtr.Zero)
                        {
                            PasteEngine.Log("clip '" + formatName + "': GlobalLock=NULL");
                            return null;
                        }
                        try
                        {
                            // SIZE_T GlobalSize(HGLOBAL) - size is the RETURN value
                            int len = (int)GlobalSize(h);
                            PasteEngine.Log("clip '" + formatName + "': handle ok, size=" + len);
                            if (len <= 0) return null;
                            byte[] data = new byte[len];
                            Marshal.Copy(ptr, data, 0, len);
                            return data;
                        }
                        finally { GlobalUnlock(h); }
                    }
                    finally { CloseClipboard(); }
                }
                PasteEngine.Log("clip '" + formatName + "': OpenClipboard busy, attempt " + attempt);
                System.Threading.Thread.Sleep(50);
            }
            return null;
        }

        // enumerate all formats currently on the clipboard (diagnostics)
        public static string DescribeFormats()
        {
            if (!OpenClipboard(IntPtr.Zero)) return "(clipboard busy)";
            try
            {
                List<string> names = new List<string>();
                uint f = 0;
                while (true)
                {
                    f = EnumClipboardFormats(f) ;
                    if (f == 0) break;
                    if (f == 1) names.Add("CF_TEXT");
                    else if (f == 13) names.Add("CF_UNICODETEXT");
                    else if (f == 15) names.Add("CF_LOCALE");
                    else
                    {
                        var sb = new System.Text.StringBuilder(256);
                        int n = GetClipboardFormatNameW(f, sb, 256);
                        names.Add(n > 0 ? sb.ToString() : ("#" + f));
                    }
                }
                return names.Count == 0 ? "(empty)" : string.Join(", ", names);
            }
            finally { CloseClipboard(); }
        }

        public static string GetUnicodeText()
        {
            if (!OpenClipboard(IntPtr.Zero)) return null;
            try
            {
                IntPtr h = GetClipboardData(13 /* CF_UNICODETEXT */);
                if (h == IntPtr.Zero) return null;
                IntPtr ptr = GlobalLock(h);
                if (ptr == IntPtr.Zero) return null;
                try { return Marshal.PtrToStringUni(ptr); }
                finally { GlobalUnlock(h); }
            }
            finally { CloseClipboard(); }
        }

        // ---------- image reading (2026-11-18) ----------

        [DllImport("user32.dll")] public static extern IntPtr GetClipboardOwner();
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        // read raw bytes for a numeric clipboard format (CF_DIB etc.)
        public static byte[] GetFormatDataById(uint fmt)
        {
            for (int attempt = 0; attempt < 5; attempt++)
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    try
                    {
                        IntPtr h = GetClipboardData(fmt);
                        if (h == IntPtr.Zero) return null;
                        IntPtr ptr = GlobalLock(h);
                        if (ptr == IntPtr.Zero) return null;
                        try
                        {
                            int len = (int)GlobalSize(h);
                            if (len <= 0) return null;
                            byte[] data = new byte[len];
                            Marshal.Copy(ptr, data, 0, len);
                            return data;
                        }
                        finally { GlobalUnlock(h); }
                    }
                    finally { CloseClipboard(); }
                }
                System.Threading.Thread.Sleep(50);
            }
            return null;
        }

        // Chromium puts ready-made PNG bytes under the registered "PNG" format;
        // most other sources only provide CF_DIB, which we wrap into a BMP file
        // and re-encode to PNG via System.Drawing.
        public static byte[] GetImageData()
        {
            // some writers use delayed rendering: the handle can be NULL on the
            // first try while the owner renders on demand, so retry a few times
            byte[] png = null;
            for (int attempt = 0; attempt < 3 && png == null; attempt++)
            {
                png = GetFormatData("PNG");
                if (png == null && attempt < 2)
                {
                    PasteEngine.Log("image: 'PNG' format empty, retry " + (attempt + 1));
                    System.Threading.Thread.Sleep(60);
                }
            }
            if (png != null && png.Length > 0)
            {
                PasteEngine.Log("image: 'PNG' format, " + png.Length + " bytes");
                return png;
            }
            byte[] dib = GetFormatDataById(8 /* CF_DIB */);
            if (dib == null || dib.Length == 0)
                dib = GetFormatDataById(17 /* CF_DIBV5, same BITMAPINFO-style layout */);
            if (dib != null && dib.Length > 0)
            {
                PasteEngine.Log("image: CF_DIB(V5), " + dib.Length + " bytes -> converting to PNG");
                byte[] conv = DibToPng(dib);
                if (conv != null) PasteEngine.Log("image: DIB converted, " + conv.Length + " PNG bytes");
                return conv;
            }
            return null;
        }

        // Convert a CF_DIB (BITMAPINFOHEADER + palette + pixels, no file header)
        // to PNG by prepending a 14-byte BITMAPFILEHEADER and decoding as BMP.
        public static byte[] DibToPng(byte[] dib)
        {
            if (dib == null || dib.Length < 40) return null;
            try
            {
                int headerSize = BitConverter.ToInt32(dib, 0);
                short bitCount = BitConverter.ToInt16(dib, 14);
                int clrUsed = BitConverter.ToInt32(dib, 32);
                int paletteEntries = clrUsed > 0 ? clrUsed : (bitCount <= 8 ? (1 << bitCount) : 0);
                int offBits = 14 + headerSize + paletteEntries * 4;
                if (offBits <= 0 || offBits >= 14 + dib.Length) return null;
                byte[] bmp = new byte[14 + dib.Length];
                bmp[0] = (byte)'B';
                bmp[1] = (byte)'M';
                Array.Copy(BitConverter.GetBytes(bmp.Length), 0, bmp, 2, 4);
                Array.Copy(BitConverter.GetBytes(offBits), 0, bmp, 10, 4);
                Array.Copy(dib, 0, bmp, 14, dib.Length);
                using (MemoryStream ms = new MemoryStream(bmp))
                using (System.Drawing.Bitmap bm = new System.Drawing.Bitmap(ms))
                using (MemoryStream outMs = new MemoryStream())
                {
                    bm.Save(outMs, System.Drawing.Imaging.ImageFormat.Png);
                    return outMs.ToArray();
                }
            }
            catch (Exception ex)
            {
                PasteEngine.Log("DIB->PNG conversion failed: " + ex.Message);
                return null;
            }
        }

        [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr hWnd, StringBuilder sb, int max);

        // window TITLE of the clipboard owner's top-level window: for Electron apps
        // and browsers this carries the page/conversation name ("对话 - WorkBuddy",
        // "页面标题 - Google Chrome"), which is far more specific than a process name.
        public static string OwnerWindowTitle()
        {
            try
            {
                IntPtr h = GetClipboardOwner();
                if (h == IntPtr.Zero) return null;
                IntPtr root = GetAncestor(h, 2 /* GA_ROOT */);
                if (root == IntPtr.Zero) root = h;
                var sb = new StringBuilder(512);
                int n = GetWindowTextW(root, sb, 512);
                string t = n > 0 ? sb.ToString().Trim() : null;
                PasteEngine.Log("clipboard owner window title: " + (t ?? "(none)"));
                return t;
            }
            catch (Exception ex)
            {
                PasteEngine.Log("window title probe failed: " + ex.Message);
                return null;
            }
        }

        // which application put the current content on the clipboard
        public static string OwnerProcessName()
        {            try
            {
                IntPtr h = GetClipboardOwner();
                if (h == IntPtr.Zero) return null;
                uint pid;
                GetWindowThreadProcessId(h, out pid);
                if (pid == 0) return null;
                using (Process p = Process.GetProcessById((int)pid))
                {
                    string n = p.ProcessName;
                    PasteEngine.Log("clipboard owner: pid=" + pid + " process=" + n);
                    return n;
                }
            }
            catch (Exception ex)
            {
                PasteEngine.Log("clipboard owner probe failed: " + ex.Message);
                return null;
            }
        }
    }

    // ---------- pure IDispatch invoker (no ITypeInfo / typelib needed) ----------
    // Rationale: this machine's OneNote typelib is not registered; the CLR late binder
    // (Type.InvokeMember / dynamic) throws TYPE_E_LIBNOTREGISTERED. PowerShell's binder
    // falls back to raw GetIDsOfNames/Invoke - which is exactly what we do here.

    public class OutParam
    {
        public string Value;
    }

    internal static class Dispatch
    {
        [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDispatchApi
        {
            [PreserveSig] int GetTypeInfoCount(out int pctinfo);
            [PreserveSig] int GetTypeInfo(int iTInfo, uint lcid, out IntPtr ppTInfo);
            [PreserveSig] int GetIDsOfNames(ref Guid riid,
                [In, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgszNames,
                int cNames, uint lcid,
                [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
            [PreserveSig] int Invoke(int dispIdMember, ref Guid riid, uint lcid, ushort wFlags,
                ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
                [Out, MarshalAs(UnmanagedType.Struct)] out object pVarResult,
                IntPtr pExcepInfo, out int puArgErr);
        }

        private const ushort DISPATCH_METHOD = 1;
        private const ushort DISPATCH_PROPERTYGET = 2;
        private const ushort VT_BYREF_BSTR = 0x4008;

        // real IApplication dispinterface IID (from ONENOTE.EXE\3 typelib; PIA's
        // {A47D3223} is NOT the real IID and can never be QI'd)
        private static readonly Guid IApplicationDispIid = new Guid("452AC71A-B655-4967-A208-A4CC39DD7949");
        private static System.Runtime.InteropServices.ComTypes.ITypeInfo _appTypeInfo;   // cached working typeinfo

        // ITypeInfo (IID verified: oaidl.h standard). Methods declared in fixed vtable order;
        // only the slots up to Invoke are declared (server never calls back into us).
        [ComImport, Guid("00020429-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ITypeInfoApi
        {
            [PreserveSig] int GetTypeAttr(out IntPtr ppTypeAttr);
            [PreserveSig] int GetTypeComp(out IntPtr ppTComp);
            [PreserveSig] int GetFuncDesc(uint index, out IntPtr ppFuncDesc);
            [PreserveSig] int GetVarDesc(uint index, out IntPtr ppVarDesc);
            [PreserveSig] int GetNames(int memid, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] rgBstrNames, int cMaxNames, out int pcNames);
            [PreserveSig] int GetRefTypeOfImplType(uint index, out IntPtr href);
            [PreserveSig] int GetImplTypeFlags(uint index, out int pImplTypeFlags);
            [PreserveSig] int GetIDsOfNames(IntPtr rgszNames, int cNames, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pMemId);
            [PreserveSig] int Invoke(IntPtr pvInstance, int memid, ushort wFlags,
                ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
                IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr);
        }

        // VARIANT as written by GetNativeVariantForObject on x64 (24 bytes)
        private static readonly int VariantSize = 24;

        public static object GetProperty(object target, string name)
        {
            return InvokeOne(target, name, null, DISPATCH_PROPERTYGET);
        }

        /// <summary>Invoke a method. OutParam slots receive the returned out strings.</summary>
        public static void CallMethod(object target, string name, object[] args)
        {
            InvokeOne(target, name, args, DISPATCH_METHOD);
        }

        private static object InvokeOne(object target, string name, object[] args, ushort flags)
        {
            IDispatchApi d = (IDispatchApi)target;
            Guid empty = Guid.Empty;
            int[] dispIds = new int[1];
            int hr = d.GetIDsOfNames(ref empty, new string[] { name }, 1, 0, dispIds);
            PasteEngine.Log("resolve '" + name + "': GetIDsOfNames hr=0x" + hr.ToString("X8") +
                (hr >= 0 ? " dispid=0x" + dispIds[0].ToString("X") : ""));

            int n = args == null ? 0 : args.Length;
            IntPtr varMem = IntPtr.Zero;
            List<IntPtr> memToFree = new List<IntPtr>();   // co-task memory
            List<IntPtr> bstrToFree = new List<IntPtr>();  // BSTRs we allocated
            List<IntPtr> byrefSlots = new List<IntPtr>();  // BSTR slot pointers for OutParam
            List<int> byrefIndex = new List<int>();        // their position in DISPPARAMS (reversed)
            IntPtr[] before = null;
            try
            {
                if (n > 0)
                {
                    varMem = Marshal.AllocCoTaskMem(n * VariantSize);
                    memToFree.Add(varMem);
                    // DISPPARAMS rgvarg[0] is the LAST argument -> reverse order
                    for (int i = 0; i < n; i++)
                    {
                        object a = args[n - 1 - i];
                        IntPtr p = varMem + i * VariantSize;
                        OutParam op = a as OutParam;
                        if (op != null)
                        {
                            // VT_BYREF|VT_BSTR: slot holds pointer to BSTR
                            // 2026-11-05 OWNERSHIP FIX: do NOT pre-register this BSTR for
                            // freeing. Cross-process, the oleaut proxy may SysFreeString our
                            // original BSTR and write a fresh one into the slot when copying
                            // the [out] value back (DllHost heap-corruption crash 0xc0000374
                            // after GetPageContent - double free). Rule: free exactly the
                            // FINAL slot pointer, once, in the finally block below.
                            IntPtr slot = Marshal.AllocCoTaskMem(IntPtr.Size);
                            memToFree.Add(slot);
                            IntPtr bstr = Marshal.StringToBSTR("");
                            Marshal.WriteIntPtr(slot, bstr);
                            byrefSlots.Add(slot);
                            byrefIndex.Add(i);
                            WriteVariant(p, VT_BYREF_BSTR, slot);
                        }
                        else
                        {
                            Marshal.GetNativeVariantForObject(a ?? DBNull.Value, p);
                        }
                    }
                }

                System.Runtime.InteropServices.ComTypes.DISPPARAMS dp =
                    new System.Runtime.InteropServices.ComTypes.DISPPARAMS();
                dp.rgvarg = varMem;
                dp.rgdispidNamedArgs = IntPtr.Zero;
                dp.cArgs = n;
                dp.cNamedArgs = 0;

                // snapshot out-slot BSTR pointers so we can detect whether the server
                // actually wrote the [out] values (raw Invoke can return S_OK and still
                // not write - seen out-of-proc on this machine)
                before = new IntPtr[byrefSlots.Count];
                for (int k = 0; k < byrefSlots.Count; k++)
                    before[k] = Marshal.ReadIntPtr(byrefSlots[k]);

                object result = null;
                int argErr;
                bool done = false;
                if (hr >= 0)
                {
                    foreach (ushort wf in new ushort[] { flags, (ushort)(flags | DISPATCH_PROPERTYGET) })
                    {
                        hr = d.Invoke(dispIds[0], ref empty, 0, wf, ref dp, out result, IntPtr.Zero, out argErr);
                        PasteEngine.Log("raw Invoke '" + name + "' wFlags=" + wf + ": hr=0x" + hr.ToString("X8") +
                            (result != null ? " result=" + result.GetType().Name : ""));
                        if (hr >= 0) { done = true; break; }
                    }
                }

                // verify out slots were written; if S_OK-but-empty, fall through to ITypeInfo
                bool outOk = true;
                if (byrefSlots.Count > 0)
                {
                    outOk = false;
                    for (int k = 0; k < byrefSlots.Count; k++)
                    {
                        IntPtr slot = byrefSlots[k];
                        IntPtr bstr = Marshal.ReadIntPtr(slot);
                        if (bstr != IntPtr.Zero && bstr != before[k]) { outOk = true; break; }
                        string s = (bstr == IntPtr.Zero) ? null : Marshal.PtrToStringBSTR(bstr);
                        if (!string.IsNullOrEmpty(s)) { outOk = true; break; }
                    }
                    PasteEngine.Log("out slots written after raw Invoke: " + outOk);
                }

                if (done && outOk)
                {
                    ReadBackOutParams(byrefSlots, bstrToFree, byrefIndex, n, args);
                    return result;
                }

                // ---------- ITypeInfo fallback (typelib-driven marshaling) ----------
                object tiResult;
                bool tiOk = InvokeViaTypeInfo(target, name, flags, ref dp, byrefSlots.Count > 0, out tiResult);
                if (tiOk)
                {
                    ReadBackOutParams(byrefSlots, bstrToFree, byrefIndex, n, args);
                    return tiResult;
                }
                if (done)
                {
                    // raw S_OK but out empty, ITypeInfo also failed
                    ReadBackOutParams(byrefSlots, bstrToFree, byrefIndex, n, args);
                    return result;
                }
                throw new COMException("Invoke('" + name + "') failed hr=0x" + hr.ToString("X8"), hr);
            }
            finally
            {
                // free each out-slot's FINAL BSTR pointer exactly once (see ownership fix
                // above). Covers paths where ReadBackOutParams never ran.
                foreach (IntPtr slot in byrefSlots)
                {
                    IntPtr b = (slot != IntPtr.Zero) ? Marshal.ReadIntPtr(slot) : IntPtr.Zero;
                    if (b != IntPtr.Zero && !bstrToFree.Contains(b)) bstrToFree.Add(b);
                }
                foreach (IntPtr b in bstrToFree) { if (b != IntPtr.Zero) Marshal.FreeBSTR(b); }
                // release BSTRs/references inside non-byref arg variants before freeing
                if (varMem != IntPtr.Zero)
                {
                    for (int i = 0; i < n; i++)
                    {
                        if (byrefIndex.Contains(i)) continue;
                        CleanupVariant(varMem + i * VariantSize);
                    }
                }
                foreach (IntPtr p in memToFree) { if (p != IntPtr.Zero) Marshal.FreeCoTaskMem(p); }
            }
        }

        private static void CleanupVariant(IntPtr p)
        {
            short vt = Marshal.ReadInt16(p, 0);
            IntPtr payload = Marshal.ReadIntPtr(p, 8);
            if (payload == IntPtr.Zero) return;
            if (vt == 8 /*VT_BSTR*/) Marshal.FreeBSTR(payload);
            else if (vt == 9 /*VT_DISPATCH*/ || vt == 13 /*VT_UNKNOWN*/)
                Marshal.Release(payload);
        }

        private static void ReadBackOutParams(List<IntPtr> byrefSlots, List<IntPtr> bstrToFree,
            List<int> byrefIndex, int n, object[] args)
        {
            for (int k = 0; k < byrefSlots.Count; k++)
            {
                IntPtr slot = byrefSlots[k];
                IntPtr bstr = Marshal.ReadIntPtr(slot);
                string val = (bstr == IntPtr.Zero) ? null : Marshal.PtrToStringBSTR(bstr);
                if (bstr != IntPtr.Zero && !bstrToFree.Contains(bstr))
                    bstrToFree.Add(bstr);   // server-allocated; free it
                int origIndex = n - 1 - byrefIndex[k];
                OutParam op = args[origIndex] as OutParam;
                if (op != null)
                {
                    op.Value = val;
                    PasteEngine.Log("out param #" + origIndex + " length=" + (val == null ? -1 : val.Length));
                }
            }
        }

        // ---------- ITypeInfo fallback ----------

        private static bool InvokeViaTypeInfo(object target, string name, ushort flags,
            ref System.Runtime.InteropServices.ComTypes.DISPPARAMS dp, bool hasOut, out object result)
        {
            result = null;
            System.Runtime.InteropServices.ComTypes.ITypeInfo ti = GetAppTypeInfo(target);
            if (ti == null) return false;
            int memid = FindMemidByName(ti, name);
            if (memid == -1) return false;

            IntPtr varRes = Marshal.AllocCoTaskMem(VariantSize);
            WriteVariant(varRes, 0 /*VT_EMPTY*/, IntPtr.Zero);
            IntPtr pUnk = Marshal.GetIUnknownForObject(target);
            int argErr;
            try
            {
                ti.Invoke(pUnk, memid, (short)flags, ref dp, varRes, IntPtr.Zero, out argErr);
                result = Marshal.GetObjectForNativeVariant(varRes);
                PasteEngine.Log("ITypeInfo.Invoke '" + name + "' OK" +
                    (result != null ? " result=" + result.GetType().Name : ""));
                _appTypeInfo = ti;   // remember the working typeinfo
                return true;
            }
            catch (Exception ex)
            {
                PasteEngine.Log("ITypeInfo.Invoke '" + name + "' failed: " + ex.Message);
                return false;
            }
            finally
            {
                Marshal.Release(pUnk);
                Marshal.FreeCoTaskMem(varRes);
            }
        }

        private static int FindMemidByName(System.Runtime.InteropServices.ComTypes.ITypeInfo ti, string name)
        {
            IntPtr pAttr;
            ti.GetTypeAttr(out pAttr);
            var attr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)
                Marshal.PtrToStructure(pAttr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));
            int memid = -1;
            for (short f = 0; f < attr.cFuncs; f++)
            {
                IntPtr pFd;
                ti.GetFuncDesc(f, out pFd);
                var fd = (System.Runtime.InteropServices.ComTypes.FUNCDESC)
                    Marshal.PtrToStructure(pFd, typeof(System.Runtime.InteropServices.ComTypes.FUNCDESC));
                string[] names = new string[1]; int pc;
                ti.GetNames(fd.memid, names, 1, out pc);
                if (names.Length > 0 && names[0] == name) memid = fd.memid;
                ti.ReleaseFuncDesc(pFd);
                if (memid != -1) break;
            }
            ti.ReleaseTypeAttr(pAttr);
            return memid;
        }

        private static System.Runtime.InteropServices.ComTypes.ITypeInfo GetAppTypeInfo(object target)
        {
            if (_appTypeInfo != null) return _appTypeInfo;

            // candidate 1: whatever GetTypeInfo(0,0) hands us (out-of-proc this returned a
            // typeinfo whose Invoke was E_NOTIMPL; in-proc may differ)
            List<System.Runtime.InteropServices.ComTypes.ITypeInfo> candidates =
                new List<System.Runtime.InteropServices.ComTypes.ITypeInfo>();
            try
            {
                IDispatchApi d = (IDispatchApi)target;
                IntPtr tiPtr;
                int hr = d.GetTypeInfo(0, 0, out tiPtr);
                PasteEngine.Log("GetTypeInfo(0,0) hr=0x" + hr.ToString("X8"));
                if (hr >= 0)
                    candidates.Add((System.Runtime.InteropServices.ComTypes.ITypeInfo)
                        Marshal.GetObjectForIUnknown(tiPtr));
            }
            catch (Exception ex) { PasteEngine.Log("GetTypeInfo(0,0) exception: " + ex.Message); }

            // candidate 2: the IApplication dispinterface from ONENOTE.EXE\3
            try
            {
                string exe = Process.GetCurrentProcess().MainModule.FileName;
                string tlbPath = exe + @"\3";
                System.Runtime.InteropServices.ComTypes.ITypeLib tlb;
                int hr = LoadTypeLibEx(tlbPath, 2 /*REGKIND_NONE*/, out tlb);
                PasteEngine.Log("LoadTypeLibEx('" + tlbPath + "') hr=0x" + hr.ToString("X8"));
                if (hr >= 0)
                {
                    int count = tlb.GetTypeInfoCount();
                    for (int i = 0; i < count; i++)
                    {
                        System.Runtime.InteropServices.ComTypes.ITypeInfo ti;
                        tlb.GetTypeInfo(i, out ti);
                        IntPtr pAttr;
                        ti.GetTypeAttr(out pAttr);
                        var attr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)
                            Marshal.PtrToStructure(pAttr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));
                        bool match = attr.typekind == System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_DISPATCH
                            && attr.guid == IApplicationDispIid;
                        ti.ReleaseTypeAttr(pAttr);
                        if (match) { candidates.Add(ti); break; }
                    }
                }
            }
            catch (Exception ex) { PasteEngine.Log("LoadTypeLibEx fallback exception: " + ex.Message); }

            if (candidates.Count == 0) return null;
            return candidates[0];
        }

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int LoadTypeLibEx(string szFile, int regkind,
            [MarshalAs(UnmanagedType.Interface)] out System.Runtime.InteropServices.ComTypes.ITypeLib pptlib);

        private static void WriteVariant(IntPtr p, ushort vt, IntPtr data)
        {
            Marshal.WriteInt16(p, 0, (short)vt);      // vt
            Marshal.WriteInt16(p, 2, 0);              // wReserved1
            Marshal.WriteInt16(p, 4, 0);              // wReserved2
            Marshal.WriteInt16(p, 6, 0);              // wReserved3
            Marshal.WriteIntPtr(p, 8, data);          // union
            Marshal.WriteIntPtr(p, 16, IntPtr.Zero);  // (alignment pad on x64)
        }
    }

    // ---------- OneNote late-bound COM wrapper (pure IDispatch) ----------

    // ---------- OneNote wrapper (pure late binding via IDispatch) ----------
    // 2026-11-05: PIA IApplication cast FAILS in-proc (QI {A47D3223} -> E_NOINTERFACE,
    // addin.log 15:59/16:16). The host object is a pure dispinterface (real IID
    // {452AC71A}). All calls go through the Dispatch invoker; signatures verified from
    // the ONENOTE.EXE\3 typelib (GetHierarchy/GetPageContent/UpdatePageContent = 4
    // params each, cParamsOpt=0 -> schema arg passed explicitly).
    // Enum values verified from the PIA: hsPages=4, piSelection=2, xs2013=2.

    public class OneNoteApp
    {
        private readonly object _app;
        private OneNoteApp(object app) { _app = app; }

        public static OneNoteApp Resolve(object appObj, bool allowCreate)
        {
            if (appObj != null)
                return new OneNoteApp(appObj);
            try
            {
                return new OneNoteApp(Marshal.GetActiveObject("OneNote.Application"));
            }
            catch (Exception ex)
            {
                PasteEngine.Log("GetActiveObject failed: " + ex.Message);
                // 2026-11-06: modern OneNote does NOT register in the ROT - GetActiveObject
                // has failed with MK_E_UNAVAILABLE on EVERY hotkey run (see addin.log), even
                // with OneNote open. CoCreateInstance binds to the running singleton though.
                // So for the hotkey path (allowCreate=false): refuse ONLY when no ONENOTE
                // process exists - otherwise CoCreate would spawn a headless instance.
                if (!allowCreate)
                {
                    bool running = false;
                    try { running = Process.GetProcessesByName("ONENOTE").Length > 0; }
                    catch (Exception pex) { PasteEngine.Log("process probe failed: " + pex.Message); }
                    if (!running)
                    {
                        PasteEngine.Log("no ONENOTE.EXE process - not creating a headless instance");
                        return null;
                    }
                    PasteEngine.Log("ONENOTE.EXE running - falling back to CoCreate (binds to running instance)");
                }
            }
            try
            {
                Type t = Type.GetTypeFromProgID("OneNote.Application");
                if (t == null) return null;
                return new OneNoteApp(Activator.CreateInstance(t));
            }
            catch (Exception ex) { PasteEngine.Log("CoCreate failed: " + ex.Message); return null; }
        }

        public string CurrentPageId()
        {
            try
            {
                object windows = Dispatch.GetProperty(_app, "Windows");
                if (windows == null) { PasteEngine.Log("Windows property returned null"); return null; }
                object curWin = Dispatch.GetProperty(windows, "CurrentWindow");
                if (curWin == null) { PasteEngine.Log("CurrentWindow property returned null"); return null; }
                object pageId = Dispatch.GetProperty(curWin, "CurrentPageId");
                return pageId as string;
            }
            catch (Exception ex) { PasteEngine.Log("CurrentPageId failed: " + ex.Message); return null; }
        }

        public string GetHierarchy()
        {
            OutParam xml = new OutParam();
            Dispatch.CallMethod(_app, "GetHierarchy", new object[] { "", 4 /*hsPages*/, xml, 2 /*xs2013*/ });
            return xml.Value;
        }

        public string GetPageContent(string pageId)
        {
            OutParam xml = new OutParam();
            Dispatch.CallMethod(_app, "GetPageContent", new object[] { pageId, xml, 2 /*piSelection*/, 2 /*xs2013*/ });
            return xml.Value;
        }

        public void UpdatePageContent(string xml)
        {
            Dispatch.CallMethod(_app, "UpdatePageContent",
                new object[] { xml, new DateTime(), 2 /*xs2013*/, false });
        }

        // scratch-page support for schema validation (probe harness)
        public string CreateNewPage(string sectionId)
        {
            OutParam pageId = new OutParam();
            Dispatch.CallMethod(_app, "CreateNewPage", new object[] { sectionId, pageId, 0 /*npsDefault*/ });
            return pageId.Value;
        }

        public void DeleteHierarchy(string objectId)
        {
            Dispatch.CallMethod(_app, "DeleteHierarchy",
                new object[] { objectId, new DateTime(), true /*deletePermanently*/ });
        }

        // wrap an externally supplied host object (About dialog uses _appObj)
        public static OneNoteApp Wrap(object appObj)
        {
            return appObj == null ? null : new OneNoteApp(appObj);
        }

        public string HostVersion()
        {
            try { return Dispatch.GetProperty(_app, "Version") as string; }
            catch { return null; }
        }
    }

    // ---------- conversion engine (ported from md2onexml.ps1 v4) ----------

    public class MdItem
    {
        public char Kind;        // 'p' | 'li' | 'h' | 't' (table) | 'i' (image)
        public int Level;        // list nesting level (li), -1 for p/h/t
        public string ListType;  // 'ul' | 'ol'
        public string Text;
        public double Pad;
        public string Style;
        public List<List<string>> Rows;   // 't': table rows, each cell holds inline HTML
        public List<List<string>> CellColors;  // 't': parallel to Rows, per-cell background color ("#RRGGBB" or "")
        public string Data;      // 'i': base64-encoded PNG payload (2026-11-18)
        public double ImgW;      // 'i': display width in points (px * 0.75)
        public double ImgH;      // 'i': display height in points
    }

    public static class PasteEngine
    {
        private static readonly List<MdItem> Items = new List<MdItem>();

        private class Cur
        {
            public char Kind;
            public int Level;
            public string ListType;
            public double Pad;
            public string Style;
            public string Text = "";
            public bool Flushed;
        }

        private static Cur _cur;
        private static int _listDepth;
        private static bool _sawNested;
        private static bool _inPre;
        private static List<string> _stack;
        private static List<double> _padStack;
        private static int _liDepth;                  // >0 while inside <li>...</li>
        private static List<string> _skipStack;      // non-content subtrees (svg/script/style/...)

        // table capture state (2026-11-06: HTML <table> -> native one:Table)
        private static bool _inTable;
        private static int _tableDepth;                    // nested-table guard
        private static List<List<string>> _rows;
        private static List<string> _row;
        private static bool _inCell;
        private static bool _cellTh;
        private static System.Text.StringBuilder _cell;
        private static string _cellColor;
        private static List<List<string>> _rowColors;   // parallel to _rows
        private static List<string> _curRowColors;      // parallel to _row
        private static int _imgPlaceholders;            // <img src> placeholders awaiting fetch
        private static int _svgDepth;                   // >0 while collecting an inline <svg>
        private static System.Text.StringBuilder _svgBuf;
        private static int _svgPlaceholders;

        public static string LogPath
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MD2OneNote", "addin.log");
            }
        }

        public static void Log(string msg)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                File.AppendAllText(LogPath,
                    "[" + DateTime.Now.ToString("HH:mm:ss.fff") + "] " + msg + "\r\n", Encoding.UTF8);
            }
            catch { }
        }

        // ============ entry point ============

        public static void Run(object appObj, string htmlFileOverride, bool allowCreate = true)
        {
            Log("run start (app=" + (appObj != null ? "attached" : "active-object") +
                (htmlFileOverride != null ? ", testfile" : "") +
                (allowCreate ? "" : ", no-create") + ")");

            OneNoteApp app = OneNoteApp.Resolve(appObj, allowCreate);
            if (app == null)
            {
                Error("OneNote not reachable. Open OneNote first.");
                return;
            }

            // ---------- read clipboard ----------
            string htmlFmt = "";
            string md = "";
            try
            {
                if (!string.IsNullOrEmpty(htmlFileOverride) && File.Exists(htmlFileOverride))
                {
                    byte[] raw = File.ReadAllBytes(htmlFileOverride);
                    if (raw != null && raw.Length > 0)
                        htmlFmt = Encoding.UTF8.GetString(raw);
                }
                else
                {
                    byte[] raw = RawClipboard.GetFormatData("HTML Format");
                    if (raw != null && raw.Length > 0)
                        htmlFmt = Encoding.UTF8.GetString(raw);   // CF_HTML is UTF-8 by spec
                    md = RawClipboard.GetUnicodeText() ?? "";
                }
            }
            catch (Exception ex) { Log("clipboard read exception: " + ex.Message); }

            Log("clipboard read: htmlFmt=" + (htmlFmt == null ? 0 : htmlFmt.Length) + " chars, text=" + (md == null ? 0 : md.Length) + " chars");
            Log("clipboard formats: " + RawClipboard.DescribeFormats());
            if (htmlFmt == null || htmlFmt.Length == 0)
            {
                string snip = (md ?? "").Replace("\r", "\\r").Replace("\n", "\\n");
                if (snip.Length > 120) snip = snip.Substring(0, 120) + "...";
                Log("text head: " + snip);
            }

            // image / source-owner / source-URL / window title: real paste path only
            byte[] img = null;
            string owner = null, srcUrl = null, winTitle = null;
            if (string.IsNullOrEmpty(htmlFileOverride))
            {
                try { img = RawClipboard.GetImageData(); } catch (Exception ex) { Log("image read failed: " + ex.Message); }
                try { owner = RawClipboard.OwnerProcessName(); } catch (Exception ex) { Log("owner probe failed: " + ex.Message); }
                try { winTitle = RawClipboard.OwnerWindowTitle(); } catch (Exception ex) { Log("window title probe failed: " + ex.Message); }
                srcUrl = ExtractSourceUrl(htmlFmt);
                if (srcUrl != null) Log("source URL: " + srcUrl);
                Log("extras: img=" + (img != null ? img.Length + "B" : "none") + ", owner=" + (owner ?? "none") +
                    ", title=" + (winTitle ?? "none") + ", url=" + (srcUrl != null ? "yes" : "no"));
            }

            List<MdItem> items = null;
            if (!string.IsNullOrWhiteSpace(htmlFmt))
            {
                // embed payloads live INSIDE <script> data, which ExtractFragment
                // strips - so extract them from the RAW CF_HTML first
                List<string> embeds = ExtractEmbedPayloads(htmlFmt);
                string html = ExtractFragment(htmlFmt);
                items = ParseHtmlToItems(html);
                foreach (string p in embeds)
                    items.Add(new MdItem { Kind = 'S', Level = -1, ListType = "", Text = p });
                if (items.Count == 0 && !string.IsNullOrWhiteSpace(md))
                    items = ParseMarkdownToItems(md);
                ResolvePlaceholders(items);   // fetch <img> srcs, rasterize inline svg
                Log("parsed html: items=" + items.Count);
            }
            else if (!string.IsNullOrWhiteSpace(md))
            {
                items = ParseMarkdownToItems(md);
                Log("parsed markdown: items=" + items.Count);
            }
            if (items == null) items = new List<MdItem>();

            if (items.Count == 0 && img == null)
            {
                Error("Clipboard is empty. Copy content first, then paste.");
                return;
            }

            AppendExtras(items, img, owner, srcUrl, winTitle);
            AssignOrderedNumbers(items);
            string outlineBody = EmitOutline(items);

            InsertIntoPage(app, outlineBody);
        }

        private static void Error(string msg)
        {
            Log("ERROR: " + msg);
            if (Environment.GetEnvironmentVariable("MD2ONENOTE_NOUI") == "1") return; // silent mode for tests
            MessageBoxW(IntPtr.Zero, msg + "\r\n\r\n(log: " + LogPath + ")", "MD2OneNote", 0x30);
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

        // ============ CF_HTML fragment extraction ============

        public static string ExtractFragment(string htmlFmt)
        {
            string html = htmlFmt;
            Match sf = Regex.Match(htmlFmt, @"StartFragment:(\d+)");
            Match ef = Regex.Match(htmlFmt, @"EndFragment:(\d+)");
            if (sf.Success && ef.Success)
            {
                byte[] bytes = Encoding.UTF8.GetBytes(htmlFmt);
                int s = int.Parse(sf.Groups[1].Value, CultureInfo.InvariantCulture);
                int e = int.Parse(ef.Groups[1].Value, CultureInfo.InvariantCulture);
                if (e > s && e <= bytes.Length)
                    html = Encoding.UTF8.GetString(bytes, s, e - s);
            }
            html = Regex.Replace(html, @"(?s)<!--.*?-->", "");
            html = Regex.Replace(html, @"(?s)<style.*?</style>", "");
            html = Regex.Replace(html, @"(?s)<script.*?</script>", "");
            return html;
        }

        // ============ HTML -> items (path A) ============

        // source URL from the CF_HTML header ("SourceURL:<...>" line, set by browsers etc.)
        public static string ExtractSourceUrl(string htmlFmt)
        {
            if (string.IsNullOrEmpty(htmlFmt)) return null;
            Match m = Regex.Match(htmlFmt, @"(?m)^SourceURL:(.+?)\s*$");
            if (!m.Success) return null;
            string url = m.Groups[1].Value.Trim();
            return url.Length > 0 ? url : null;
        }

        public static string EscapeText(string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        public static string Cdata(string inner)
        {
            inner = inner.Replace("]]>", "]]]]><![CDATA[>");
            return "<![CDATA[" + inner + "]]>";
        }

        // ---------- image + source-line helpers (2026-11-18) ----------

        // public passthrough: DIB->PNG is a pure conversion (RawClipboard is internal)
        public static byte[] DibToPng(byte[] dib) { return RawClipboard.DibToPng(dib); }

        public static bool PngDimensions(byte[] png, out int w, out int h)
        {
            w = 0; h = 0;
            if (png == null || png.Length < 24) return false;
            if (png[0] != 0x89 || png[1] != 0x50 || png[2] != 0x4E || png[3] != 0x47) return false;
            w = (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19];
            h = (png[20] << 24) | (png[21] << 16) | (png[22] << 8) | png[23];
            return w > 0 && h > 0;
        }

        public static MdItem MakeImageItem(byte[] png)
        {
            if (png == null || png.Length == 0) return null;
            int w, h;
            if (!PngDimensions(png, out w, out h))
            {
                Log("image: not a valid PNG (no IHDR), skipped");
                return null;
            }
            // px -> points at the standard 96dpi mapping (same convention as GetPadPt)
            return new MdItem
            {
                Kind = 'i', Level = -1, ListType = "",
                Data = Convert.ToBase64String(png),
                ImgW = Math.Max(1, w * 0.75),
                ImgH = Math.Max(1, h * 0.75)
            };
        }

        // trailing source line: which app the content came from, its window/page
        // title, and (only for real web links) the URL.
        // URL only shown when it is a real web link (http/https); file:// and
        // other schemes are dropped (user request 2026-11-18 11:13)
        public static MdItem BuildSourceItem(string owner, string url, string winTitle)
        {
            string t = "";
            if (!string.IsNullOrEmpty(owner)) t += "\u6765\u81EA " + EscapeText(owner);
            string title = TrimAppSuffix(winTitle);
            if (!string.IsNullOrEmpty(title))
            {
                if (t.Length > 0) t += " \u00B7 ";
                t += "\u300C" + EscapeText(title) + "\u300D";
            }
            if (!string.IsNullOrEmpty(url) &&
                (url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                 url.StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
            {
                if (t.Length > 0) t += " \u00B7 ";
                // href keeps the raw URL (clickable, faithful); display text is a
                // readable translation (percent-decoded, query dropped, shortened)
                t += "<a href=\"" + EscapeText(url) + "\">" + EscapeText(UrlDisplayText(url)) + "</a>";
            }
            if (t.Length == 0) return null;
            return new MdItem { Kind = 'p', Level = -1, ListType = "", Text = t, Pad = 0.0, Style = "" };
        }

        // browsers/Electron append the app name to the window title
        // ("页面标题 - Google Chrome"); strip it to keep just the page name
        public static string TrimAppSuffix(string title)
        {
            if (string.IsNullOrEmpty(title)) return null;
            string t = title.Trim();
            Match m = Regex.Match(t,
                "\\s+-\\s+(Google Chrome|Microsoft Edge|Chromium|Firefox|WorkBuddy|Chrome|Edge|QQ\\u6D4F\\u89C8\\u5668|360\\u5B89\\u5168\\u6D4F\\u89C8\\u5668|123pan|App)$",
                RegexOptions.IgnoreCase);
            if (m.Success) t = t.Substring(0, m.Index).Trim();
            return t.Length > 0 ? t : null;
        }

        // readable form of a source URL for the display text only:
        // percent-encoding decoded (repeatedly, %257B -> %7B -> {), query and
        // fragment dropped, file:// shown as a plain Windows path, middle-truncated
        public static string UrlDisplayText(string url)
        {
            if (string.IsNullOrEmpty(url)) return url;
            string shown = url;
            for (int i = 0; i < 3; i++)
            {
                string dec;
                try { dec = Uri.UnescapeDataString(shown); }
                catch { break; }
                if (dec == shown || dec.Length == 0 || dec.IndexOf('\0') >= 0) break;
                shown = dec;
            }
            if (shown.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
            {
                string p = shown;
                int hash = p.IndexOf('#'); if (hash >= 0) p = p.Substring(0, hash);
                int q = p.IndexOf('?'); if (q >= 0) p = p.Substring(0, q);
                p = Regex.Replace(p, "^file://(localhost)?/?", "", RegexOptions.IgnoreCase);
                p = p.Replace('/', '\\');
                return TruncateMiddle(p, 72);
            }
            Match m = Regex.Match(shown, "^(?i:https?)://([^/?#]+)(/[^?#]*)?");
            if (m.Success)
            {
                string s = m.Groups[1].Value;
                string path = m.Groups[2].Value;
                if (!string.IsNullOrEmpty(path))
                {
                    path = path.TrimEnd('/');
                    if (path.Length > 0) s += path;
                }
                return TruncateMiddle(s, 72);
            }
            return TruncateMiddle(shown, 72);
        }

        private static string TruncateMiddle(string s, int max)
        {
            if (string.IsNullOrEmpty(s) || s.Length <= max) return s;
            int head = max * 2 / 3;
            return s.Substring(0, head) + "\u2026" + s.Substring(s.Length - (max - head));
        }

        private static void AppendExtras(List<MdItem> items, byte[] imgPng, string owner, string srcUrl, string winTitle)
        {
            if (imgPng != null)
            {
                MdItem im = MakeImageItem(imgPng);
                if (im != null)
                {
                    items.Add(im);
                    Log("image item appended: " + im.ImgW + "x" + im.ImgH + " pt, " + imgPng.Length + " bytes");
                }
            }
            MdItem src = BuildSourceItem(owner, srcUrl, winTitle);
            if (src != null) { items.Add(src); Log("source line appended"); }
        }

        // ---------- <img src> resolution: data:/file:/http(s) -> PNG bytes ----------

        private static void FinishSvg()
        {
            string svg = _svgBuf.ToString();
            _svgBuf = null;
            if (svg.Length > 40)   // ignore tiny stray fragments
            {
                Items.Add(new MdItem { Kind = 'S', Level = -1, ListType = "", Text = svg });
                _svgPlaceholders++;
                Log("svg captured: " + svg.Length + " chars");
            }
        }

        // replace 'I' (img src) and 'S' (inline svg) placeholders with real image
        // items (or drop/keep-text them). Called from Run() at paste time ONLY —
        // never inside ParseHtmlToItems, so parser tests stay pure and IO-free.
        public static void ResolvePlaceholders(List<MdItem> items)
        {
            int found = 0, ok = 0;
            for (int i = 0; i < items.Count; i++)
            {
                if (items[i].Kind == 'S')
                {
                    found++;
                    string svg = items[i].Text;
                    byte[] png = RasterizeSvg(svg);
                    if (png != null)
                    {
                        MdItem im = MakeImageItem(png);
                        if (im != null)
                        {
                            items[i] = im;
                            ok++;
                            Log("svg rasterized: " + im.ImgW + "x" + im.ImgH + " pt, " + png.Length + " bytes");
                            // "paste the picture AND the script": content diagrams
                            // carry <title>/<desc> - keep them as a caption line
                            string cap = SvgCaption(svg);
                            if (cap.Length > 0)
                            {
                                items.Insert(i + 1, new MdItem
                                {
                                    Kind = 'p', Level = -1, ListType = "",
                                    Text = "[\u56FE\u6CE8] " + EscapeText(cap),
                                    Pad = 0.0, Style = ""
                                });
                                i++;   // skip the inserted caption
                            }
                            continue;
                        }
                    }
                    // fallback: keep the diagram's text content, never silently lose it
                    string txt = SvgTextContent(svg);
                    if (txt.Length > 0)
                    {
                        items[i] = new MdItem
                        {
                            Kind = 'p', Level = -1, ListType = "",
                            Text = "[\u56FE\u8868\u65E0\u6CD5\u6E32\u67D3\uFF0C\u4FDD\u7559\u6587\u5B57] " + EscapeText(txt),
                            Pad = 0.0, Style = ""
                        };
                        ok++;
                        Log("svg rasterize failed, text fallback kept (" + txt.Length + " chars)");
                    }
                    else
                    {
                        Log("svg rasterize failed and no text content, dropped");
                        items.RemoveAt(i);
                        i--;
                    }
                    continue;
                }
                if (items[i].Kind != 'I') continue;
                found++;
                string src = items[i].Text;
                byte[] png2 = FetchImageAsPng(src);
                if (png2 != null)
                {
                    MdItem im2 = MakeImageItem(png2);
                    if (im2 != null)
                    {
                        items[i] = im2;
                        ok++;
                        Log("img resolved: " + im2.ImgW + "x" + im2.ImgH + " pt from " + ShortSrc(src));
                        continue;
                    }
                }
                Log("img dropped (fetch/decode failed): " + ShortSrc(src));
                items.RemoveAt(i);
                i--;
            }
            if (found > 0) Log("html images: found=" + found + " resolved=" + ok);
        }

        // rasterize inline SVG to PNG via a headless Edge/Chromium screenshot.
        // Verified path: msedge --headless --screenshot with an isolated profile.
        public static byte[] RasterizeSvg(string svg)
        {
            string browser = FindHeadlessBrowser();
            if (browser == null) { Log("svg rasterize: no msedge/chrome found"); return null; }
            int w = SvgDim(svg, "width", 800);
            int h = SvgDim(svg, "height", 600);
            w = Math.Max(40, Math.Min(2400, w));
            h = Math.Max(40, Math.Min(2400, h));
            string work = Path.Combine(Path.GetTempPath(), "MD2OneNote", "svg");
            try { Directory.CreateDirectory(work); } catch { return null; }
            string htmlPath = Path.Combine(work, "r-" + Guid.NewGuid().ToString("N") + ".html");
            string pngPath = Path.Combine(work, "r.png");
            string profPath = Path.Combine(work, "profile");
            try
            {
                File.WriteAllText(htmlPath,
                    "<html><body style=\"margin:0;padding:0;background:#FFFFFF\">" + svg + "</body></html>",
                    new System.Text.UTF8Encoding(false));
                try { File.Delete(pngPath); } catch { }
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = browser,
                    Arguments = "--headless --disable-gpu --no-first-run --hide-scrollbars " +
                                "--user-data-dir=\"" + profPath + "\" " +
                                "--screenshot=\"" + pngPath + "\" " +
                                "--window-size=" + w + "," + h + " " +
                                "--default-background-color=FFFFFFFF " +
                                "\"file:///" + htmlPath.Replace('\\', '/') + "\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (System.Diagnostics.Process p = System.Diagnostics.Process.Start(psi))
                {
                    // the screenshot file is the real completion signal; the browser
                    // process may linger, so poll the file and kill afterwards
                    byte[] png = null;
                    for (int i = 0; i < 75; i++)   // up to 15 s
                    {
                        System.Threading.Thread.Sleep(200);
                        if (File.Exists(pngPath))
                        {
                            try
                            {
                                byte[] cur = File.ReadAllBytes(pngPath);
                                if (cur.Length > 8 && cur[0] == 0x89 && cur[1] == 0x50) { png = cur; break; }
                            }
                            catch { }
                        }
                    }
                    try { if (!p.HasExited) p.Kill(); } catch { }
                    if (png == null) { Log("svg rasterize: screenshot never appeared"); return null; }
                    int pw, ph;
                    PngDimensions(png, out pw, out ph);
                    Log("svg rasterize: window " + w + "x" + h + " -> png " + pw + "x" + ph);
                    return png;
                }
            }
            catch (Exception ex)
            {
                Log("svg rasterize exception: " + ex.Message);
                return null;
            }
            finally
            {
                try { File.Delete(htmlPath); } catch { }
            }
        }

        private static string FindHeadlessBrowser()
        {
            string[] candidates =
            {
                @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                @"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                @"C:\Program Files\Google\Chrome\Application\chrome.exe",
                @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            };
            foreach (string c in candidates) if (File.Exists(c)) return c;
            return null;
        }

        // SVG width/height attribute (px) or viewBox; fallback to def.
        // width="100%" and other non-pixel values are ignored -> viewBox wins.
        private static int SvgDim(string svg, string attr, int def)
        {
            Match m = Regex.Match(svg, "<svg[^>]*?\\b" + attr + "\\s*=\\s*\"([^\"]*)\"",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (m.Success)
            {
                string v = m.Groups[1].Value.Trim().TrimEnd('p', 'x');
                double num;
                if (v.Length > 0 && !v.Contains("%") &&
                    double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out num) && num > 0)
                    return (int)Math.Round(num);
            }
            m = Regex.Match(svg, "<svg[^>]*?\\bviewBox\\s*=\\s*\"[\\d.]+\\s+[\\d.]+\\s+([\\d.]+)\\s+([\\d.]+)\"",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (m.Success)
            {
                double v = attr == "width"
                    ? double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture)
                    : double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
                if (v > 0) return (int)Math.Round(v);
            }
            return def;
        }

        // all visible text inside the SVG (fallback when rasterization is impossible)
        private static string SvgTextContent(string svg)
        {
            var parts = new List<string>();
            foreach (Match m in Regex.Matches(svg, "<(?:text|tspan)\\b[^>]*>(.*?)</(?:text|tspan)>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline))
            {
                string t = System.Net.WebUtility.HtmlDecode(Regex.Replace(m.Groups[1].Value, "<[^>]+>", "")).Trim();
                if (t.Length > 0) parts.Add(t);
            }
            return string.Join("; ", parts);
        }

        public static byte[] FetchImageAsPng(string src)
        {
            if (string.IsNullOrEmpty(src)) return null;
            byte[] raw = null;
            if (src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase))
            {
                int comma = src.IndexOf(',');
                if (comma > 0 && src.IndexOf("base64", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    try { raw = Convert.FromBase64String(src.Substring(comma + 1)); }
                    catch (Exception ex) { Log("img data-uri decode failed: " + ex.Message); }
                }
            }
            else if (src.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
            {
                string p = Uri.UnescapeDataString(src);
                p = Regex.Replace(p, "^file://(localhost)?/?", "", RegexOptions.IgnoreCase).Replace('/', '\\');
                try { if (File.Exists(p)) raw = File.ReadAllBytes(p); }
                catch (Exception ex) { Log("img file read failed: " + ex.Message + " | " + p); }
            }
            else if (src.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                     src.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                raw = HttpGetBytes(src);
            }
            if (raw == null || raw.Length == 0) return null;
            return DecodeImageToPng(raw);
        }

        // PNG passes through; everything System.Drawing can read (JPEG/GIF/BMP/TIFF)
        // is re-encoded to PNG so one:Image only ever carries format="png"
        public static byte[] DecodeImageToPng(byte[] raw)
        {
            if (raw == null || raw.Length == 0) return null;
            int w, h;
            if (PngDimensions(raw, out w, out h)) return raw;
            try
            {
                using (MemoryStream ms = new MemoryStream(raw))
                using (System.Drawing.Bitmap bm = new System.Drawing.Bitmap(ms))
                using (MemoryStream outMs = new MemoryStream())
                {
                    bm.Save(outMs, System.Drawing.Imaging.ImageFormat.Png);
                    return outMs.ToArray();
                }
            }
            catch (Exception ex)
            {
                Log("image decode failed: " + ex.Message);
                return null;
            }
        }

        private static byte[] HttpGetBytes(string url)
        {
            try
            {
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";
                req.Timeout = 6000;
                req.ReadWriteTimeout = 6000;
                req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) 2Note/1.0";
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                {
                    if (resp.ContentType != null && resp.ContentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase))
                    {
                        Log("img http rejected (content-type=" + resp.ContentType + ")");
                        return null;
                    }
                    using (MemoryStream ms = new MemoryStream())
                    {
                        resp.GetResponseStream().CopyTo(ms);
                        return ms.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                Log("img http failed: " + ex.Message + " | " + ShortSrc(url));
                return null;
            }
        }

        private static string ShortSrc(string s)
        {
            if (s == null) return "";
            return s.Length > 80 ? s.Substring(0, 80) + "\u2026" : s;
        }

        // decorative icon detection: aria-hidden or tiny viewport (WorkBuddy chat
        // serializes 16x16 button icons as inline svg - rendering those as "images"
        // would produce three junk pictures per copied message)
        private static bool IsIconSvg(string attrs)
        {
            if (attrs == null) return false;
            if (attrs.IndexOf("aria-hidden", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            foreach (string dim in new string[] { "width", "height" })
            {
                Match m = Regex.Match(attrs, "\\b" + dim + "\\s*=\\s*\"([^\"]*)\"", RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    string v = m.Groups[1].Value.Trim();
                    double num;
                    if (v.EndsWith("px")) v = v.Substring(0, v.Length - 2).Trim();
                    if (double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out num) && num <= 24.0)
                        return true;
                }
            }
            return false;
        }

        // title + desc of a content diagram, kept as a caption line under the image
        private static string SvgCaption(string svg)
        {
            string title = TagInner(svg, "title");
            string desc = TagInner(svg, "desc");
            string cap = "";
            if (title.Length > 0) cap += title;
            if (desc.Length > 0) cap += (cap.Length > 0 ? "\uFF1A" : "") + desc;
            return cap.Trim();
        }

        private static string TagInner(string svg, string tag)
        {
            Match m = Regex.Match(svg, "<" + tag + "\\b[^>]*>(.*?)</" + tag + ">",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!m.Success) return "";
            return System.Net.WebUtility.HtmlDecode(Regex.Replace(m.Groups[1].Value, "<[^>]+>", "")).Trim();
        }

        // ---------- WorkBuddy chat iframe payloads ----------
        // Chat diagram attachments ("widget" cards) keep their whole iframe document
        // (styles + bootstrap script) inside a DOM ATTRIBUTE. The diagram itself is
        // NOT in the copied HTML as <img>/<svg> - it lives as a JS object literal:
        //   var initialContent = {"html":"\\u003csvg role=\\&quot;img\\&quot; ..."}
        // where the attribute serializer doubled every backslash and turned every
        // quote into &quot;. Verified end-to-end against a real capture (11-18):
        // decode chain = HtmlDecode -> undo doubled backslashes -> JS-unescape.
        public static List<string> ExtractEmbedPayloads(string html)
        {
            var result = new List<string>();
            var seen = new HashSet<string>();
            if (string.IsNullOrEmpty(html)) return result;
            int idx = html.IndexOf("initialContent", StringComparison.Ordinal);
            while (idx >= 0 && idx < html.Length)
            {
                int max = Math.Min(400000, html.Length - idx);
                string seg = System.Net.WebUtility.HtmlDecode(html.Substring(idx, max));
                int a = seg.IndexOf('{');
                // real payloads nest a theme object -> "}}"; allow a flat object too
                int b = a >= 0 ? seg.IndexOf("}};", a, StringComparison.Ordinal) : -1;
                if (b < 0 && a >= 0) b = seg.IndexOf("};", a, StringComparison.Ordinal);
                if (a >= 0 && b > a)
                {
                    // }}; form: b+2 covers both closing braces; }; form: b+2 covers
                    // the brace and the semicolon, harmless outside the string scan
                    string objText = seg.Substring(a, b - a + 2);
                    string payload = ExtractHtmlValue(UndoDoubledBackslashes(objText));
                    if (payload != null && payload.TrimStart().StartsWith("<svg", StringComparison.OrdinalIgnoreCase)
                        && payload.Contains("</svg>") && seen.Add(payload))
                    {
                        result.Add(payload);
                        Log("embed payload extracted: " + payload.Length + " chars");
                    }
                }
                int next = html.IndexOf("initialContent", idx + 14, StringComparison.Ordinal);
                if (next == idx) break;
                idx = next;
            }
            return result;
        }

        // every '\' in the attribute was serialized as '\\' (verified byte-level);
        // a lone '\' is kept as-is just in case
        private static string UndoDoubledBackslashes(string s)
        {
            var sb = new System.Text.StringBuilder(s.Length);
            int i = 0;
            while (i < s.Length)
            {
                if (s[i] == '\\' && i + 1 < s.Length && s[i + 1] == '\\') { sb.Append('\\'); i += 2; }
                else { sb.Append(s[i]); i++; }
            }
            return sb.ToString();
        }

        // read the JS string value of "html":"..." with lenient JS unescaping
        // (\\n, \\t, \\uXXXX, \\" \\' \\/ \\\\ - NOT strict JSON: \\' is legal JS)
        private static string ExtractHtmlValue(string objText)
        {
            Match m = Regex.Match(objText, "\"html\"\\s*:\\s*\"");
            if (!m.Success) return null;
            var outSb = new System.Text.StringBuilder();
            int i = m.Index + m.Length;
            while (i < objText.Length)
            {
                char c = objText[i];
                if (c == '"') break;
                if (c == '\\')
                {
                    if (i + 1 >= objText.Length) break;
                    char nc = objText[i + 1];
                    switch (nc)
                    {
                        case 'n': outSb.Append('\n'); i += 2; break;
                        case 't': outSb.Append('\t'); i += 2; break;
                        case 'r': outSb.Append('\r'); i += 2; break;
                        case 'u':
                            if (i + 6 <= objText.Length)
                            {
                                int code;
                                if (int.TryParse(objText.Substring(i + 2, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
                                {
                                    outSb.Append((char)code); i += 6; break;
                                }
                            }
                            outSb.Append(c); i++; break;
                        default: outSb.Append(nc); i += 2; break;   // \" \' \\ \/ etc
                    }
                    continue;
                }
                outSb.Append(c); i++;
            }
            string v = outSb.ToString();
            return v.Length > 0 ? v : null;
        }

        private static double GetPadPt(string attrs)
        {
            Match m = Regex.Match(attrs, @"padding-left\s*:\s*([\d.]+)\s*(em|px|pt)");
            if (!m.Success) m = Regex.Match(attrs, @"margin-left\s*:\s*([\d.]+)\s*(em|px|pt)");
            if (!m.Success) return 0.0;
            double v = double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            switch (m.Groups[2].Value)
            {
                case "em": return v * 12.0;
                case "px": return v * 0.75;
                default: return v;
            }
        }

        private static void FlushCur()
        {
            if (_cur == null) return;
            if (!_cur.Flushed)
            {
                string t = _cur.Text.Trim();
                if (t.Length > 0 || _cur.Kind == 'l')
                {
                    Items.Add(new MdItem { Kind = _cur.Kind, Level = _cur.Level, ListType = _cur.ListType, Text = t, Pad = _cur.Pad, Style = _cur.Style });
                }
                _cur.Flushed = true;
            }
            // CRITICAL: clear the slot. Keeping a flushed Cur here made every text
            // node that arrives outside a recognised block tag (e.g. Chromium/Qwen
            // wraps plain paragraphs in <div> instead of <p>) get appended to an
            // item that had already been emitted -> silently lost content.
            _cur = null;
        }

        private static void StartCur(char kind, int level, string lt, double pad, string style)
        {
            if (_cur != null && !_cur.Flushed) FlushCur();
            _cur = new Cur { Kind = kind, Level = level, ListType = lt, Pad = pad, Style = style, Text = "", Flushed = false };
        }

        private static void AppendInline(string s)
        {
            if (_inCell)
            {
                if (_cell != null) _cell.Append(s);
                return;
            }
            if (_cur == null) StartCur('p', -1, "", 0.0, "");
            _cur.Text += s;
        }

        // ============ table capture helpers ============

        private static void EndCell(bool th)
        {
            if (_cell != null && _row != null)
            {
                string c = _cell.ToString().Trim();
                if (th && c.Length > 0) c = "<b>" + c + "</b>";
                _row.Add(c);
                if (_curRowColors != null) _curRowColors.Add(_cellColor ?? "");
            }
            _inCell = false;
            _cellTh = false;
            _cell = null;
            _cellColor = null;
        }

        private static void EndRow()
        {
            if (_inCell) EndCell(_cellTh);
            if (_row != null && _row.Count > 0 && _rows != null)
            {
                _rows.Add(_row);
                if (_rowColors != null && _curRowColors != null) _rowColors.Add(_curRowColors);
            }
            _row = null;
            _curRowColors = null;
            _inCell = false;
        }

        private static void EndTable()
        {
            EndRow();
            if (_rows != null && _rows.Count > 0)
            {
                FlushCur();
                Items.Add(new MdItem
                {
                    Kind = 't', Level = -1, ListType = "",
                    Text = "", Pad = 0.0, Style = "",
                    Rows = new List<List<string>>(_rows),
                    CellColors = (_rowColors != null) ? new List<List<string>>(_rowColors) : null
                });
            }
            _inTable = false;
            _tableDepth = 0;
            _rows = null;
            _row = null;
            _rowColors = null;
            _curRowColors = null;
            _inCell = false;
            _cell = null;
        }

        // extract cell background color from a <td>/<th> open tag; "" when absent.
        // 2026-11-06: verified one:Cell accepts shadingColor="#RRGGBB" (round-trips).
        // STRICT output contract: only "#RRGGBB" or "" ever leave this function.
        // Anything else (rgb strings, transparent/url, truncated 4-digit hex) once
        // produced invalid shadingColor values and OneNote rejected the whole
        // UpdatePageContent with DISP_E_EXCEPTION (addin.log 16:58).
        public static string ParseBgColor(string attrs)
        {
            if (attrs == null || attrs.Length == 0) return "";
            string v = "";
            Match m = Regex.Match(attrs, @"bgcolor\s*=\s*[""']?(#?[0-9a-fA-F]{8}|#?[0-9a-fA-F]{6}|#?[0-9a-fA-F]{3})(?![0-9a-fA-F])[""']?");
            if (m.Success)
                v = m.Groups[1].Value;
            else
            {
                m = Regex.Match(attrs, @"background(?:-color)?\s*:\s*(#[0-9a-fA-F]{8}|#[0-9a-fA-F]{6}|#[0-9a-fA-F]{3}(?![0-9a-fA-F])|rgb\([^)]+\)|[a-zA-Z]+)");
                if (!m.Success) return "";
                v = m.Groups[1].Value;
            }

            // rgb(r,g,b) / rgba(...) -> take the first three components
            if (v.StartsWith("rgb", StringComparison.OrdinalIgnoreCase))
            {
                Match rm = Regex.Match(v, @"(\d+)\D+(\d+)\D+(\d+)");
                if (!rm.Success) return "";
                return "#" + byte.Parse(rm.Groups[1].Value).ToString("X2") +
                       byte.Parse(rm.Groups[2].Value).ToString("X2") +
                       byte.Parse(rm.Groups[3].Value).ToString("X2");
            }

            // named colors: whitelist only (transparent/url/unknown -> "")
            string lower = v.ToLowerInvariant();
            switch (lower)
            {
                case "red": return "#FF0000";
                case "yellow": return "#FFFF00";
                case "orange": return "#FFA500";
                case "gold": return "#FFD700";
                case "green": return "#008000";
                case "lime": return "#00FF00";
                case "blue": return "#0000FF";
                case "cyan": case "aqua": return "#00FFFF";
                case "purple": return "#800080";
                case "pink": return "#FFC0CB";
                case "gray": case "grey": return "#808080";
                case "silver": return "#C0C0C0";
                case "white": return "#FFFFFF";
                case "black": return "#000000";
                case "lightgray": case "lightgrey": return "#D3D3D3";
                case "lightyellow": return "#FFFFE0";
                case "lightblue": return "#ADD8E6";
                case "lightgreen": return "#90EE90";
                case "lightred": return "#FF7F7F";
            }

            // hex path (# optional in the bgcolor form)
            if (v.StartsWith("#")) v = v.Substring(1);
            if (v.Length == 8) v = v.Substring(0, 6);   // drop alpha channel
            if (v.Length == 3) v = new string(new char[] { v[0], v[0], v[1], v[1], v[2], v[2] });
            if (v.Length != 6 || !Regex.IsMatch(v, "^[0-9a-fA-F]{6}$")) return "";
            return "#" + v.ToUpperInvariant();
        }

        // ============ HTML tokenizer / block classification ============

        // Quote-aware tag scanner. Regex.Split(html, "(<[^>]+>)") breaks as soon as
        // an attribute value contains '>' (inline CSS, data-* payloads) and also
        // swallows stray '<' characters. Real clipboard HTML (Chromium, OneNote)
        // contains newlines inside tags and very long inline style attributes, so
        // scan characters and honour quoted regions instead.
        private static List<string> Tokenize(string html)
        {
            List<string> toks = new List<string>();
            StringBuilder buf = new StringBuilder();
            int i = 0, n = html.Length;
            while (i < n)
            {
                if (html[i] == '<' && i + 1 < n &&
                    (char.IsLetter(html[i + 1]) || html[i + 1] == '/' ||
                     html[i + 1] == '!' || html[i + 1] == '?'))
                {
                    int j = i + 1;
                    char quote = '\0';
                    bool end = false;
                    while (j < n)
                    {
                        char c = html[j];
                        if (quote != '\0') { if (c == quote) quote = '\0'; }
                        else if (c == '"' || c == '\'') quote = c;
                        else if (c == '>') { end = true; break; }
                        j++;
                    }
                    if (end)
                    {
                        if (buf.Length > 0) { toks.Add(buf.ToString()); buf.Length = 0; }
                        toks.Add(html.Substring(i, j - i + 1));
                        i = j + 1;
                        continue;
                    }
                }
                buf.Append(html[i]);
                i++;
            }
            if (buf.Length > 0) toks.Add(buf.ToString());
            return toks;
        }

        private static bool IsBlockTag(string tag)
        {
            switch (tag)
            {
                // Chromium/Qwen chat renders EVERY plain paragraph as <div>, not <p>.
                // Treat generic containers as paragraph boundaries on both edges.
                case "div": case "section": case "article": case "header": case "footer":
                case "main": case "aside": case "nav": case "figure": case "figcaption":
                case "center": case "address": case "details": case "summary":
                case "dl": case "dt": case "dd": case "fieldset": case "form":
                case "hgroup": case "hr":
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsSkipTag(string tag)
        {
            switch (tag)
            {
                // NOTE: svg is NOT skipped since 2026-11-18: chat-rendered diagrams
                // (mermaid etc.) arrive as inline <svg>; we capture and rasterize them
                case "script": case "style": case "head": case "noscript":
                case "iframe": case "canvas": case "math":
                case "template": case "select": case "object": case "embed":
                case "audio": case "video": case "map":
                    return true;
                default:
                    return false;
            }
        }

        public static List<MdItem> ParseHtmlToItems(string html)
        {
            Items.Clear();
            _listDepth = 0;
            _sawNested = false;
            _inPre = false;
            _cur = null;
            _liDepth = 0;
            _skipStack = new List<string>();
            _stack = new List<string>();
            _padStack = new List<double>();
            _inTable = false;
            _tableDepth = 0;
            _rows = null;
            _row = null;
            _inCell = false;
            _cellTh = false;
            _cell = null;
            _cellColor = null;
            _rowColors = null;
            _curRowColors = null;
            _imgPlaceholders = 0;
            _svgDepth = 0;
            _svgBuf = null;
            _svgPlaceholders = 0;

            List<string> tokens = Tokenize(html);
            foreach (string tok in tokens)
            {
                if (tok.Length == 0) continue;

                // ---- inline SVG capture: accumulate tokens verbatim until the
                // matching </svg>, then park the whole element as a placeholder
                if (_svgDepth > 0)
                {
                    Match sv = Regex.Match(tok, "^<\\s*(/?)([a-zA-Z0-9:_\\-]+)(.*)>$", RegexOptions.Singleline);
                    if (sv.Success && sv.Groups[2].Value.ToLowerInvariant() == "svg")
                    {
                        _svgBuf.Append(tok);
                        if (sv.Groups[1].Value == "/")
                        {
                            _svgDepth--;
                            if (_svgDepth == 0) FinishSvg();
                        }
                        else if (sv.Groups[3].Value.TrimEnd().EndsWith("/"))
                        {
                            _svgDepth--;   // self-closing <svg/>
                            if (_svgDepth == 0) FinishSvg();
                        }
                        else _svgDepth++;
                    }
                    else _svgBuf.Append(tok);
                    continue;
                }

                if (tok.Length > 1 && tok[0] == '<' && tok[tok.Length - 1] == '>')
                {
                    // Tokenizer already found the real closing '>' outside quotes,
                    // so attrs may itself contain '>' (inline CSS / viewBox / data-*).
                    Match tm = Regex.Match(tok, @"^<\s*(/?)([a-zA-Z0-9:_\-]+)(.*)>$", RegexOptions.Singleline);
                    if (!tm.Success) continue;
                    bool closing = tm.Groups[1].Value == "/";
                    string tag = tm.Groups[2].Value.ToLowerInvariant();
                    string attrs = tm.Groups[3].Value;

                    // Skip non-content subtrees (icon <svg>, leftover script/style).
                    if (IsSkipTag(tag))
                    {
                        if (closing)
                        {
                            if (_skipStack.Count > 0 && _skipStack[_skipStack.Count - 1] == tag)
                                _skipStack.RemoveAt(_skipStack.Count - 1);
                        }
                        // stack-top dedup: serialized script DATA often contains
                        // literal "<script>" strings (JS templates); a second push
                        // made the single real </script> leave stray entries and
                        // silently swallow ALL content after it (WorkBuddy evidence)
                        else if (_skipStack.Count == 0 || _skipStack[_skipStack.Count - 1] != tag)
                            _skipStack.Add(tag);
                        if (!_inCell && !_inTable) FlushCur();
                        continue;
                    }
                    // The closer of a skipped icon svg must POP even though svg is
                    // no longer in IsSkipTag - otherwise the stack entry leaks and
                    // every token after it is silently swallowed.
                    if (_skipStack.Count > 0)
                    {
                        if (closing && tag == "svg" && _skipStack[_skipStack.Count - 1] == "svg")
                            _skipStack.RemoveAt(_skipStack.Count - 1);
                        continue;
                    }

                    // Track list-item scope: <div> wrappers inside an <li> must not
                    // split a single list item into several paragraphs.
                    if (tag == "li")
                    {
                        if (closing) { if (_liDepth > 0) _liDepth--; }
                        else _liDepth++;
                    }

                    // Block containers delimit paragraphs. Flushing on the opening
                    // edge matters just as much as on the closing edge: after a
                    // </p> / </h3> / </li> the current item is already emitted, so
                    // the text node inside the next <div> needs a fresh Cur.
                    if (IsBlockTag(tag) && !_inCell && !_inTable && _liDepth == 0)
                        FlushCur();

                    switch (tag)
                    {
                        case "table":
                            if (closing)
                            {
                                if (_tableDepth > 0) { _tableDepth--; }
                                else if (_inTable) { EndTable(); }
                            }
                            else if (_inTable)
                            {
                                _tableDepth++;   // nested table: treat as plain cell text
                            }
                            else
                            {
                                FlushCur();
                                _inTable = true; _tableDepth = 0;
                                _rows = new List<List<string>>();
                                _rowColors = new List<List<string>>();
                                _row = null; _curRowColors = null; _inCell = false; _cell = null;
                            }
                            break;
                        case "tr":
                            if (!_inTable || _tableDepth > 0) break;
                            if (closing) { EndRow(); }
                            else { FlushCur(); EndRow(); _row = new List<string>(); _curRowColors = new List<string>(); _inCell = false; }
                            break;
                        case "th":
                        case "td":
                            if (!_inTable || _tableDepth > 0 || _row == null) break;
                            if (closing) { EndCell(_cellTh); }
                            else
                            {
                                _inCell = true; _cellTh = (tag == "th");
                                _cell = new System.Text.StringBuilder();
                                _cellColor = ParseBgColor(attrs);
                            }
                            break;
                        case "ul":
                        case "ol":
                            if (_inCell) break;
                            if (closing)
                            {
                                FlushCur();
                                if (_listDepth > 0) _listDepth--;
                                if (_padStack.Count > 0) _padStack.RemoveAt(_padStack.Count - 1);
                                if (_stack.Count > 0) _stack.RemoveAt(_stack.Count - 1);
                            }
                            else
                            {
                                FlushCur();
                                _listDepth++;
                                if (_listDepth >= 2) _sawNested = true;
                                _stack.Add(tag);
                                _padStack.Add(GetPadPt(attrs));
                            }
                            break;
                        case "li":
                            if (_inCell) break;
                            if (closing) { FlushCur(); }
                            else
                            {
                                int lvl = Math.Max(0, _stack.Count - 1);
                                string lt = "ul";
                                if (_stack.Count > 0) lt = _stack[_stack.Count - 1];
                                double pad = 0.0;
                                if (_padStack.Count > 0) pad = _padStack[_padStack.Count - 1];
                                double lp = GetPadPt(attrs);
                                if (lp > pad) pad = lp;
                                StartCur('l', lvl, lt, pad, "");
                            }
                            break;
                        case "p": if (!_inCell) { if (closing) { FlushCur(); } else { StartCur('p', -1, "", 0.0, ""); } } break;
                        case "blockquote": if (!_inCell) { if (closing) { FlushCur(); } else { StartCur('p', -1, "", 0.0, ""); } } break;
                        case "h1": if (!_inCell) { if (closing) { FlushCur(); } else { StartCur('h', -1, "", 0.0, "font-size:20.0pt;font-weight:bold"); } } break;
                        case "h2": if (!_inCell) { if (closing) { FlushCur(); } else { StartCur('h', -1, "", 0.0, "font-size:16.0pt;font-weight:bold"); } } break;
                        case "h3":
                        case "h4":
                        case "h5":
                        case "h6": if (!_inCell) { if (closing) { FlushCur(); } else { StartCur('h', -1, "", 0.0, "font-weight:bold"); } } break;
                        case "pre": if (!_inCell) { if (closing) { _inPre = false; } else { FlushCur(); _inPre = true; } } break;
                        case "br":
                            if (_inCell) { if (_cell != null) _cell.Append(" "); }
                            else AppendInline(" ");   // never drop the separator between lines
                            break;
                        case "svg":
                            if (closing)
                            {
                                // a skipped icon svg's closer pops the skip entry
                                if (_skipStack.Count > 0 && _skipStack[_skipStack.Count - 1] == "svg")
                                    _skipStack.RemoveAt(_skipStack.Count - 1);
                                break;
                            }
                            // decorative icons (aria-hidden, tiny) are junk: skip the
                            // whole subtree exactly like the old svg-skip path
                            if (IsIconSvg(attrs)) { _skipStack.Add("svg"); break; }
                            // start capture (only at top level; cells keep text only)
                            if (_inCell || _inTable) break;
                            if (_liDepth == 0) FlushCur();
                            _svgBuf = new System.Text.StringBuilder(tok);
                            _svgDepth = attrs.TrimEnd().EndsWith("/") ? 0 : 1;
                            if (_svgDepth == 0) FinishSvg();
                            break;
                        case "img":
                            // web-copy path: images arrive as <img src> inside the HTML,
                            // NOT as clipboard bitmaps. Remember the src as a placeholder
                            // item; ResolvePlaceholders() fetches/decodes at paste time.
                            // Inside table cells we keep text only (image-in-cell deferred).
                            if (!closing && !_inCell && !_inTable)
                            {
                                Match sm = Regex.Match(attrs,
                                    "src\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)'|([^\\s>]+))",
                                    RegexOptions.IgnoreCase);
                                string src = sm.Groups[1].Value + sm.Groups[2].Value + sm.Groups[3].Value;
                                if (src.Length > 0)
                                {
                                    if (_liDepth == 0) FlushCur();
                                    Items.Add(new MdItem { Kind = 'I', Level = -1, ListType = "", Text = src });
                                    _imgPlaceholders++;
                                }
                            }
                            break;
                        case "input":
                            if (attrs.Contains("type=\"checkbox\""))
                            {
                                AppendInline(attrs.Contains("checked") ? "\u2611 " : "\u2610 ");
                            }
                            break;
                        case "b":
                        case "strong": AppendInline(closing ? "</b>" : "<b>"); break;
                        case "i":
                        case "em": AppendInline(closing ? "</i>" : "<i>"); break;
                        case "code": AppendInline(closing ? "</span>" : "<span style=\"font-family:Consolas,monospace;\">"); break;
                        case "a":
                            if (closing) { AppendInline("</a>"); }
                            else
                            {
                                Match hm = Regex.Match(attrs,
                                    "href\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)'|([^\\s>]+))",
                                    RegexOptions.IgnoreCase);
                                string href = hm.Groups[1].Value + hm.Groups[2].Value + hm.Groups[3].Value;
                                // always emit the opener: case "a" (closing) appends </a>
                                AppendInline("<a href=\"" + href + "\">");
                            }
                            break;
                        default: break;
                    }
                }
                else
                {
                    if (_skipStack.Count > 0) continue;   // text inside svg/script/style etc.
                    string txt = System.Net.WebUtility.HtmlDecode(tok);
                    if (_inPre)
                    {
                        foreach (string ln in txt.Replace("\r\n", "\n").Split('\n'))
                        {
                            if (ln.Trim().Length > 0)
                            {
                                if (_cur != null && !_cur.Flushed) FlushCur();
                                Items.Add(new MdItem
                                {
                                    Kind = 'p', Level = -1, ListType = "",
                                    Text = "<span style=\"font-family:Consolas,monospace;\">" + EscapeText(ln) + "</span>",
                                    Pad = 0.0, Style = ""
                                });
                            }
                        }
                    }
                    else
                    {
                        txt = Regex.Replace(txt, @"\s+", " ");
                        if (txt.Trim().Length == 0) continue;
                        if (_inCell) { if (_cell != null) _cell.Append(txt); continue; }
                        if (_inTable) continue;   // stray text between rows: ignore
                        if (_cur == null) StartCur('p', -1, "", 0.0, "");
                        _cur.Text += txt;
                    }
                }
            }
            FlushCur();
            // NOTE: placeholders ('I' img / 'S' svg) are NOT resolved here on purpose:
            // resolving involves network IO and headless-browser rasterization, which
            // Run() triggers explicitly via ResolvePlaceholders() at paste time only.

            bool anyList = Items.Exists(delegate(MdItem it) { return it.Kind == 'l'; });
            if (anyList && !_sawNested)
            {
                double minPad = 1e9;
                foreach (MdItem it in Items)
                    if (it.Kind == 'l' && it.Pad > 0.01 && it.Pad < minPad) minPad = it.Pad;
                if (minPad < 1e8 && minPad > 0.01)
                {
                    foreach (MdItem it in Items)
                        if (it.Kind == 'l') it.Level = (int)Math.Round(it.Pad / minPad);
                }
            }
            // return a copy so test harnesses never touch live state
            return new List<MdItem>(Items);
        }

        // ============ markdown -> items (path B) ============

        private static string ConvertInline(string s)
        {
            s = Regex.Replace(s, @"\[([^\]]+)\]\(([^)]+)\)", "<a href=\"$2\">$1</a>");
            s = Regex.Replace(s, @"\*\*([^*]+)\*\*", "<b>$1</b>");
            s = Regex.Replace(s, @"(?<!\*)\*([^*]+)\*(?!\*)", "<i>$1</i>");
            s = Regex.Replace(s, "`([^`]+)`", "<span style=\"font-family:Consolas,monospace;\">$1</span>");
            return s;
        }

        public static List<MdItem> ParseMarkdownToItems(string md)
        {
            List<MdItem> items = new List<MdItem>();
            string[] lines = md.Replace("\r\n", "\n").Split('\n');
            List<int> stack = new List<int>();
            Dictionary<int, string> types = new Dictionary<int, string>();
            bool blankBefore = false;
            int prevListLevel = -1;

            // 2026-11-06: pipe-table support (| a | b | rows)
            List<string> tbl = new List<string>();
            Action flushTbl = delegate
            {
                if (tbl.Count >= 2)
                {
                    List<List<string>> rows = new List<List<string>>();
                    bool header = false;
                    for (int i = 0; i < tbl.Count; i++)
                    {
                        string l = tbl[i].Trim();
                        string inner = l.StartsWith("|") ? l.Substring(1) : l;
                        if (inner.EndsWith("|")) inner = inner.Substring(0, inner.Length - 1);
                        string[] cells = inner.Split('|');
                        bool sep = true;
                        List<string> row = new List<string>();
                        foreach (string c0 in cells)
                        {
                            string c = c0.Trim();
                            if (c.Length > 0 && !Regex.IsMatch(c, "^:?-{3,}:?$")) sep = false;
                            row.Add(c);
                        }
                        if (sep) { header = (i == 1); continue; }
                        rows.Add(row);
                    }
                    if (rows.Count > 0)
                    {
                        for (int r = 0; r < rows.Count; r++)
                            for (int i = 0; i < rows[r].Count; i++)
                                rows[r][i] = ConvertInline(EscapeText(rows[r][i]));
                        if (header)
                            for (int i = 0; i < rows[0].Count; i++)
                                if (rows[0][i].Length > 0) rows[0][i] = "<b>" + rows[0][i] + "</b>";
                        items.Add(new MdItem { Kind = 't', Level = -1, ListType = "", Text = "", Pad = 0.0, Style = "", Rows = rows });
                    }
                }
                tbl.Clear();
            };

            foreach (string raw in lines)
            {
                string line = raw.Replace("\t", "  ");
                if (line.TrimStart().StartsWith("|"))
                {
                    tbl.Add(line);
                    blankBefore = false;
                    continue;
                }
                if (tbl.Count > 0) flushTbl();

                Match lm = Regex.Match(line, @"^(\s*)([-*+]|\d+[.)])\s+(.*)$");
                if (lm.Success)
                {
                    int indent = lm.Groups[1].Value.Length;
                    string marker = lm.Groups[2].Value;
                    string content = lm.Groups[3].Value;
                    bool isOrdered = Regex.IsMatch(marker, @"^\d");
                    if (Regex.IsMatch(content, @"^\[ \]\s*")) content = Regex.Replace(content, @"^\[ \]\s*", "\u2610 ");
                    if (Regex.IsMatch(content, @"^\[[xX]\]\s*")) content = Regex.Replace(content, @"^\[[xX]\]\s*", "\u2611 ");
                    content = ConvertInline(EscapeText(content));
                    while (stack.Count > 0 && indent < stack[stack.Count - 1]) stack.RemoveAt(stack.Count - 1);
                    if (stack.Count == 0 || indent > stack[stack.Count - 1]) stack.Add(indent);
                    int level = stack.Count - 1;
                    if (blankBefore || (prevListLevel == level && types.ContainsKey(level) && (types[level] == "ol") != isOrdered))
                        types.Remove(level);
                    if (!types.ContainsKey(level)) types[level] = isOrdered ? "ol" : "ul";
                    items.Add(new MdItem { Kind = 'l', Level = level, ListType = types[level], Text = content, Pad = 0.0, Style = "" });
                    prevListLevel = level;
                    blankBefore = false;
                }
                else if (line.Trim().Length != 0)
                {
                    stack.Clear(); types.Clear(); prevListLevel = -1; blankBefore = false;
                    string t = line.Trim();
                    Match hm = Regex.Match(t, @"^(#{1,6})\s+(.*)$");
                    if (hm.Success)
                    {
                        int h = hm.Groups[1].Value.Length;
                        string txt = ConvertInline(EscapeText(hm.Groups[2].Value));
                        string st = h == 1 ? "font-size:20.0pt;font-weight:bold" :
                                    h == 2 ? "font-size:16.0pt;font-weight:bold" : "font-weight:bold";
                        items.Add(new MdItem { Kind = 'h', Level = -1, ListType = "", Text = txt, Pad = 0.0, Style = st });
                    }
                    else
                    {
                        items.Add(new MdItem { Kind = 'p', Level = -1, ListType = "", Text = ConvertInline(EscapeText(t)), Pad = 0.0, Style = "" });
                    }
                }
                else { blankBefore = true; }
            }
            if (tbl.Count > 0) flushTbl();
            return items;
        }

        // ============ literal numbers for ordered items ============

        public static void AssignOrderedNumbers(List<MdItem> items)
        {
            Dictionary<int, int> olCounters = new Dictionary<int, int>();
            bool inOl = false;
            for (int i = 0; i < items.Count; i++)
            {
                MdItem it = items[i];
                if (it.Kind == 'l' && it.ListType == "ol")
                {
                    if (!inOl) { olCounters.Clear(); inOl = true; }
                    if (!olCounters.ContainsKey(it.Level)) olCounters[it.Level] = 0;
                    olCounters[it.Level] = olCounters[it.Level] + 1;
                    it.Text = olCounters[it.Level] + ". " + it.Text;
                }
                else { inOl = false; }
            }
        }

        // ============ XML emit ============

        private static string EmitItem(MdItem it)
        {
            if (it.Kind == 't') return EmitTable(it);
            if (it.Kind == 'I') return "";   // unresolved placeholder: never emit
            if (it.Kind == 'S') return "";   // unresolved svg placeholder: never emit
            if (it.Kind == 'i')
            {
                // structure per real-world usage of the OneNote schema: Image sits
                // inside an OE with Position + Size and a base64 one:Data payload
                return "<one:OE><one:Image format=\"png\" originalPageNumber=\"0\">" +
                       "<one:Position x=\"0\" y=\"0\" z=\"0\"/>" +
                       "<one:Size width=\"" + Num(it.ImgW) + "\" height=\"" + Num(it.ImgH) + "\"/>" +
                       "<one:Data>" + it.Data + "</one:Data></one:Image></one:OE>";
            }
            if (it.Kind == 'l')
            {
                if (it.ListType == "ul")
                    return "<one:OE><one:List><one:Bullet bullet=\"2\"/></one:List><one:T>" + Cdata(it.Text) + "</one:T>";
                return "<one:OE><one:T>" + Cdata(it.Text) + "</one:T>";
            }
            if (it.Kind == 'h')
                return "<one:OE><one:T>" + Cdata("<span style=\"" + it.Style + "\">" + it.Text + "</span>") + "</one:T>";
            return "<one:OE><one:T>" + Cdata(it.Text) + "</one:T>";
        }

        // ============ table XML emit ============
        // Element names and placement taken from OneNote's OWN table XML (dumped from
        // a live page, probe/onenote-table-ref.xml): <one:Table> sits INSIDE an
        // <one:OE> (no <one:T> sibling), contains <one:Columns> (NO count attribute)
        // then <one:Row> (NOT one:TableRow!) each with <one:Cell><one:OEChildren>.

        public static string EmitTable(MdItem it)
        {
            if (it.Rows == null || it.Rows.Count == 0) return "";
            int cols = 0;
            foreach (List<string> r in it.Rows) if (r.Count > cols) cols = r.Count;
            if (cols == 0) return "";
            bool header = it.Rows[0].Count > 0;
            foreach (string c in it.Rows[0]) if (!c.StartsWith("<b>")) header = false;
            StringBuilder sb = new StringBuilder();
            sb.Append("<one:OE><one:Table bordersVisible=\"true\"");
            if (header) sb.Append(" hasHeaderRow=\"true\"");
            sb.Append("><one:Columns>");
            for (int i = 0; i < cols; i++)
                sb.Append("<one:Column index=\"").Append(i).Append("\" width=\"120.0\"/>");
            sb.Append("</one:Columns>");
            foreach (List<string> r in it.Rows)
            {
                int rIdx = it.Rows.IndexOf(r);
                List<string> colors = (it.CellColors != null && rIdx >= 0 && rIdx < it.CellColors.Count) ? it.CellColors[rIdx] : null;
                sb.Append("<one:Row>");
                for (int i = 0; i < cols; i++)
                {
                    string c = i < r.Count ? r[i] : "";
                    string color = (colors != null && i < colors.Count) ? colors[i] : "";
                    sb.Append("<one:Cell");
                    if (color.Length > 0) sb.Append(" shadingColor=\"").Append(color).Append("\"");
                    sb.Append("><one:OEChildren><one:OE><one:T>")
                      .Append(Cdata(c))
                      .Append("</one:T></one:OE></one:OEChildren></one:Cell>");
                }
                sb.Append("</one:Row>");
            }
            sb.Append("</one:Table></one:OE>");
            return sb.ToString();
        }

        public static string EmitOutline(List<MdItem> items)
        {
            StringBuilder sb = new StringBuilder();
            int curDepth = 0;
            bool openOE = false;
            bool inList = false;

            foreach (MdItem it in items)
            {
                if (it.Kind == 'l')
                {
                    inList = true;
                    if (it.Level > curDepth)
                    {
                        sb.Append("<one:OEChildren>");
                        curDepth++;
                    }
                    else
                    {
                        while (curDepth > it.Level)
                        {
                            sb.Append("</one:OE></one:OEChildren>");
                            curDepth--;
                        }
                        if (openOE) { sb.Append("</one:OE>"); openOE = false; }
                    }
                    sb.Append(EmitItem(it));
                    openOE = true;
                }
                else
                {
                    if (inList)
                    {
                        while (curDepth > 0)
                        {
                            sb.Append("</one:OE></one:OEChildren>");
                            curDepth--;
                        }
                        inList = false;
                    }
                    if (openOE) { sb.Append("</one:OE>"); openOE = false; }
                    if (it.Kind == 't' || it.Kind == 'i')
                    {
                        // EmitTable and the Image item are SELF-CONTAINED (they open
                        // AND close their own one:OE wrapper) - leaving openOE set
                        // made the next item emit a stray </one:OE>, the whole
                        // outline was malformed and OneNote rejected EVERY degrade
                        // level (16:59 evidence: failed-update-L0..L3 all broken).
                        sb.Append(EmitItem(it));
                        openOE = false;
                    }
                    else
                    {
                        sb.Append(EmitItem(it));
                        openOE = true;
                    }
                }
            }
            while (curDepth > 0)
            {
                sb.Append("</one:OE></one:OEChildren>");
                curDepth--;
            }
            if (openOE) sb.Append("</one:OE>");

            return "<one:OEChildren>" + sb.ToString() + "</one:OEChildren>";
        }

        // ============ page location + insertion ============

        public static string Num(double d)
        {
            return d.ToString(CultureInfo.InvariantCulture);
        }

        public static void InsertIntoPage(OneNoteApp app, string outlineBody)
        {
            string pageId = app.CurrentPageId();
            if (string.IsNullOrWhiteSpace(pageId))
            {
                try
                {
                    string hier = app.GetHierarchy() ?? "";
                    Match m = Regex.Match(hier, "<one:Page\\b[^>]*isCurrentlyViewed=\"true\"[^>]*?ID=\"(\\{[^\"]+\\})\"");
                    if (!m.Success)
                        m = Regex.Match(hier, "<one:Page\\b[^>]*?ID=\"(\\{[^\"]+\\})\"[^>]*?isCurrentlyViewed=\"true\"");
                    if (m.Success) pageId = m.Groups[1].Value;
                }
                catch (Exception ex) { Log("GetHierarchy fallback failed: " + ex.Message); }
            }
            if (string.IsNullOrWhiteSpace(pageId))
            {
                Error("OneNote page not found. Open OneNote and click into a page first.");
                return;
            }
            Log("page found: " + pageId);

            string pageXml = null;
            pageXml = app.GetPageContent(pageId);
            Log("page xml: " + (pageXml == null ? 0 : pageXml.Length) + " chars");
            if (pageXml == null) { Error("Unexpected empty page XML."); return; }

            Match oeSel = Regex.Match(pageXml, "<one:OE\\b[^>]*selected=\"(?:all|partial)\"[^>]*>");
            if (oeSel.Success)
            {
                int tOpen = pageXml.IndexOf("<one:Title>");
                int tClose = pageXml.IndexOf("</one:Title>");
                if (tOpen >= 0 && tClose > tOpen && oeSel.Index > tOpen && oeSel.Index < tClose)
                    oeSel = Match.Empty;
            }

            string newXml;
            int itemsCount = CountItems(outlineBody);

            if (oeSel.Success)
            {
                int outOpen = pageXml.LastIndexOf("<one:Outline", oeSel.Index);
                int outClose = pageXml.IndexOf("</one:Outline>", oeSel.Index);
                if (outOpen >= 0 && outClose > outOpen)
                {
                    int outEnd = outClose + "</one:Outline>".Length;
                    string seg = pageXml.Substring(outOpen, Math.Min(outClose - outOpen, 3000));
                    double px = 36.0, py = -1.0, ph = 0.0;
                    Match pm = Regex.Match(seg, "<one:Position x=\"([-\\d.]+)\" y=\"([-\\d.]+)\"");
                    if (pm.Success)
                    {
                        px = double.Parse(pm.Groups[1].Value, CultureInfo.InvariantCulture);
                        py = double.Parse(pm.Groups[2].Value, CultureInfo.InvariantCulture);
                    }
                    Match sm2 = Regex.Match(seg, "<one:Size width=\"[-\\d.]+\" height=\"([-\\d.]+)\"");
                    if (sm2.Success) ph = double.Parse(sm2.Groups[1].Value, CultureInfo.InvariantCulture);

                    if (py >= 0)
                    {
                        double ny = py + ph + 15.0;
                        List<double[]> boxes = new List<double[]>(); // {y, h}
                        foreach (Match o in Regex.Matches(pageXml, "<one:Position x=\"([-\\d.]+)\" y=\"([-\\d.]+)\"[^/]*/><one:Size width=\"[\\d.]+\" height=\"([\\d.]+)\""))
                        {
                            double ox = double.Parse(o.Groups[1].Value, CultureInfo.InvariantCulture);
                            double oy = double.Parse(o.Groups[2].Value, CultureInfo.InvariantCulture);
                            double oh = double.Parse(o.Groups[3].Value, CultureInfo.InvariantCulture);
                            if (Math.Abs(ox - px) < 80) boxes.Add(new double[] { oy, oh });
                        }
                        for (int guard = 0; guard < 50; guard++)
                        {
                            bool hit = false;
                            foreach (double[] b in boxes)
                            {
                                if (ny >= (b[0] - 2) && ny < (b[0] + b[1] + 5))
                                {
                                    ny = b[0] + b[1] + 15.0;
                                    hit = true;
                                    break;
                                }
                            }
                            if (!hit) break;
                        }
                        double estH = itemsCount * 22 + 20;
                        string newOutline = "<one:Outline><one:Position x=\"" + Num(px) + "\" y=\"" + Num(ny) +
                            "\" z=\"0\"/><one:Size width=\"620.0\" height=\"" + Num(estH) + "\"/>" + outlineBody + "</one:Outline>";
                        newXml = pageXml.Substring(0, outEnd) + newOutline + pageXml.Substring(outEnd);
                        Update(app, newXml, newOutline);
                        return;
                    }
                }
            }

            int idx = pageXml.LastIndexOf("</one:Page>");
            if (idx < 0) { Error("Unexpected page XML structure."); return; }
            double estH2 = itemsCount * 22 + 20;
            string newOutline2 = "<one:Outline><one:Size width=\"620.0\" height=\"" + Num(estH2) + "\"/>" + outlineBody + "</one:Outline>";
            newXml = pageXml.Substring(0, idx) + newOutline2 + pageXml.Substring(idx);
            Update(app, newXml, newOutline2);
        }

        private static int CountItems(string outlineBody)
        {
            int n = 0;
            int i = outlineBody.IndexOf("<one:OE>", StringComparison.Ordinal);
            while (i >= 0) { n++; i = outlineBody.IndexOf("<one:OE>", i + 1, StringComparison.Ordinal); }
            return n;
        }

        // 2026-11-18: rejection no longer aborts the paste. OneNote validates the
        // WHOLE submitted page, so one bad value (e.g. an illegal shadingColor)
        // rejects everything. We now retry with progressively simpler markup -
        // OUR outline only, the user's existing page content is never touched -
        // logging every attempt, dumping each rejected XML, and telling the
        // user what actually landed and why.
        public static readonly string[] DegradeLabels = {
            "L0 full", "L1 no-shading", "L2 no-inline-style", "L3 plain-text"
        };

        public static void Update(OneNoteApp app, string newXml, string ownOutline)
        {
            string cause = null;
            for (int level = 0; level < DegradeLabels.Length; level++)
            {
                string submit = newXml;
                if (level > 0)
                {
                    string degraded = DegradeOutline(ownOutline, level);
                    if (degraded == ownOutline) continue;   // nothing left to drop here
                    // degrade notice goes INTO the content (last line), not a popup:
                    // a modal box interrupts the user mid-flow (user request 11-18)
                    degraded = AppendDegradeNotice(degraded, level, cause);
                    submit = newXml.Replace(ownOutline, degraded);
                }
                try
                {
                    app.UpdatePageContent(submit);
                    if (level == 0)
                    {
                        Log("update OK");
                    }
                    else
                    {
                        Log("update OK after degradation: " + DegradeLabels[level] +
                            " | first cause: " + cause + " | notice appended to content");
                    }
                    return;
                }
                catch (Exception ex)
                {
                    cause = ex.Message;
                    Log("update rejected at " + DegradeLabels[level] + ": " + cause);
                    DumpFailedXml(submit, level);
                }
            }
            // every level rejected: fall back to NATIVE paste. We never touched the
            // clipboard, so replaying Ctrl+V lets OneNote's own importer handle it.
            Log("all degrade levels rejected - falling back to native paste (Ctrl+V)");
            SendNativePaste();
        }

        // "[2Note 提示] 本次粘贴已降级 (...)" appended as the last line of our outline
        public static string AppendDegradeNotice(string outline, int level, string cause)
        {
            string notice = "[2Note \u63D0\u793A] \u672C\u6B21\u7C98\u8D34\u5DF2\u964D\u7EA7\uFF08" + DegradeLabels[level] +
                "\uFF09\uFF0C\u90E8\u5206\u683C\u5F0F\u88AB\u7B80\u5316\u3002\u539F\u56E0: " + (cause ?? "?");
            string item = "<one:OE><one:T>" + Cdata(notice) + "</one:T></one:OE>";
            int pos = outline.LastIndexOf("</one:OEChildren>", StringComparison.Ordinal);
            if (pos < 0) return outline + item;
            return outline.Insert(pos, item);
        }

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        private const uint KEYEVENTF_KEYUP = 0x0002;

        private static void SendNativePaste()
        {
            if (Environment.GetEnvironmentVariable("MD2ONENOTE_NOUI") == "1")
            {
                Log("native paste fallback skipped (NOUI)");
                return;
            }
            try
            {
                keybd_event(0x11 /*VK_CONTROL*/, 0, 0, UIntPtr.Zero);
                keybd_event(0x56 /*VK_V*/, 0, 0, UIntPtr.Zero);
                keybd_event(0x56, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                keybd_event(0x11, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                Log("native paste: Ctrl+V sent to foreground window");
            }
            catch (Exception ex)
            {
                Log("native paste failed: " + ex.Message);
            }
        }

        // Progressive simplification of OUR outline only.
        // level>=1: drop table cell shadingColor (the page-wide rejection cause found on 11-06)
        // level>=2: drop CSS attrs on one: elements and ALL inline markup inside CDATA
        // level>=3: tables become plain text rows, list bullets/numbers removed
        public static string DegradeOutline(string outline, int level)
        {
            string xml = outline;
            if (level >= 1)
                xml = Regex.Replace(xml, "\\s+shadingColor=\"[^\"]*\"", "");
            if (level >= 2)
            {
                xml = Regex.Replace(xml, "(<one:[A-Za-z]+\\b[^>]*?)\\s+style=\"[^\"]*\"", "$1");
                xml = StripCdataMarkup(xml);
            }
            if (level >= 3)
            {
                xml = Regex.Replace(xml, "<one:Table\\b[^>]*>.*?</one:Table>",
                    delegate(Match m) { return TableToText(m.Value); }, RegexOptions.Singleline);
                xml = Regex.Replace(xml, "<one:Image\\b[^>]*>.*?</one:Image>",
                    delegate(Match m) { return "<one:T>" + Cdata("[image omitted]") + "</one:T>"; },
                    RegexOptions.Singleline);
                xml = Regex.Replace(xml, "<one:List>.*?</one:List>", "", RegexOptions.Singleline);
            }
            return xml;
        }

        private static string StripCdataMarkup(string xml)
        {
            return Regex.Replace(xml, "<!\\[CDATA\\[(.*?)\\]\\]>",
                delegate(Match m)
                {
                    return "<![CDATA[" + Regex.Replace(m.Groups[1].Value, "<[^>]+>", "") + "]]>";
                },
                RegexOptions.Singleline);
        }

        private static string TableToText(string tableXml)
        {
            List<string> rows = new List<string>();
            foreach (Match r in Regex.Matches(tableXml, "<one:Row\\b[^>]*>(.*?)</one:Row>", RegexOptions.Singleline))
            {
                List<string> cells = new List<string>();
                foreach (Match c in Regex.Matches(r.Groups[1].Value, "<!\\[CDATA\\[(.*?)\\]\\]>", RegexOptions.Singleline))
                    cells.Add(Regex.Replace(c.Groups[1].Value, "<[^>]+>", "").Trim());
                if (cells.Count > 0) rows.Add(string.Join(" | ", cells.ToArray()));
            }
            if (rows.Count == 0) return "";
            return "<one:T>" + Cdata("[table as text] " + string.Join(" / ", rows.ToArray())) + "</one:T>";
        }

        // NotifyDegraded popup REMOVED 11-18: modal boxes interrupt the user; the
        // degrade notice is now appended to the pasted content (AppendDegradeNotice)

        // OneNote reports rejections as a bare DISP_E_EXCEPTION with no detail;
        // dump the exact submitted XML per attempt so failures can be analyzed offline.
        private static void DumpFailedXml(string xml, int level)
        {
            try
            {
                string p = Path.Combine(Path.GetDirectoryName(LogPath),
                    "failed-update-L" + level + ".xml");
                File.WriteAllText(p, xml, Encoding.UTF8);
                Log("dumped rejected xml: " + p + " (" + xml.Length + " chars)");
            }
            catch (Exception ex) { Log("dump failed: " + ex.Message); }
        }
    }
}
