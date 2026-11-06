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
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]

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
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom)
        {
            if (_hotkey != null) { _hotkey.Dispose(); _hotkey = null; }
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
                // (RCW created on the host's STA). Run(null) resolves via GetActiveObject,
                // which returns an apartment-correct proxy.
                _hotkey = new HotkeyWindow(mod, vk, delegate
                {
                    try { PasteEngine.Run(null, null); }
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
                "<button id=\"btn2NotePaste\" label=\"Paste as OneNote structure\" imageMso=\"Paste\" " +
                "onAction=\"OnPaste\" size=\"large\"/>" +
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
            if (_hwnd != IntPtr.Zero)
            {
                UnregisterHotKey(_hwnd, _id);
                _hwnd = IntPtr.Zero;
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

        public static OneNoteApp Resolve(object appObj)
        {
            if (appObj != null)
                return new OneNoteApp(appObj);
            try
            {
                return new OneNoteApp(Marshal.GetActiveObject("OneNote.Application"));
            }
            catch (Exception ex) { PasteEngine.Log("GetActiveObject failed: " + ex.Message); }
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
    }

    // ---------- conversion engine (ported from md2onexml.ps1 v4) ----------

    public class MdItem
    {
        public char Kind;        // 'p' | 'li' | 'h'
        public int Level;        // list nesting level (li), -1 for p/h
        public string ListType;  // 'ul' | 'ol'
        public string Text;
        public double Pad;
        public string Style;
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

        public static void Run(object appObj, string htmlFileOverride)
        {
            Log("run start (app=" + (appObj != null ? "attached" : "active-object") +
                (htmlFileOverride != null ? ", testfile" : "") + ")");

            OneNoteApp app = OneNoteApp.Resolve(appObj);
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

            string outlineBody;
            if (!string.IsNullOrWhiteSpace(htmlFmt))
            {
                string html = ExtractFragment(htmlFmt);
                List<MdItem> items = ParseHtmlToItems(html);
                if (items.Count == 0)
                {
                    if (string.IsNullOrWhiteSpace(md)) { Error("Clipboard has no usable content."); return; }
                    items = ParseMarkdownToItems(md);
                }
                AssignOrderedNumbers(items);
                outlineBody = EmitOutline(items);
                Log("parsed html: items=" + items.Count);
            }
            else if (!string.IsNullOrWhiteSpace(md))
            {
                List<MdItem> items = ParseMarkdownToItems(md);
                if (items.Count == 0) { Error("Clipboard is empty. Copy content first."); return; }
                AssignOrderedNumbers(items);
                outlineBody = EmitOutline(items);
                Log("parsed markdown: items=" + items.Count);
            }
            else
            {
                Error("Clipboard is empty. Copy content first, then paste.");
                return;
            }

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

        public static string EscapeText(string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        public static string Cdata(string inner)
        {
            inner = inner.Replace("]]>", "]]]]><![CDATA[>");
            return "<![CDATA[" + inner + "]]>";
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
            if (_cur != null && !_cur.Flushed)
            {
                string t = _cur.Text.Trim();
                if (t.Length > 0 || _cur.Kind == 'l')
                {
                    Items.Add(new MdItem { Kind = _cur.Kind, Level = _cur.Level, ListType = _cur.ListType, Text = t, Pad = _cur.Pad, Style = _cur.Style });
                }
                _cur.Flushed = true;
            }
        }

        private static void StartCur(char kind, int level, string lt, double pad, string style)
        {
            if (_cur != null && !_cur.Flushed) FlushCur();
            _cur = new Cur { Kind = kind, Level = level, ListType = lt, Pad = pad, Style = style, Text = "", Flushed = false };
        }

        private static void AppendInline(string s)
        {
            if (_cur == null) StartCur('p', -1, "", 0.0, "");
            _cur.Text += s;
        }

        public static List<MdItem> ParseHtmlToItems(string html)
        {
            Items.Clear();
            _listDepth = 0;
            _sawNested = false;
            _inPre = false;
            _cur = null;
            _stack = new List<string>();
            _padStack = new List<double>();

            string[] tokens = Regex.Split(html, "(<[^>]+>)");
            foreach (string tok in tokens)
            {
                if (tok.Length == 0) continue;
                if (Regex.IsMatch(tok, "^<[^>]+>$"))
                {
                    Match tm = Regex.Match(tok, @"^<(/?)([a-zA-Z0-9]+)([^>]*)>$");
                    if (!tm.Success) continue;
                    bool closing = tm.Groups[1].Value == "/";
                    string tag = tm.Groups[2].Value.ToLowerInvariant();
                    string attrs = tm.Groups[3].Value;
                    switch (tag)
                    {
                        case "ul":
                        case "ol":
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
                        case "p": if (closing) { FlushCur(); } else { StartCur('p', -1, "", 0.0, ""); } break;
                        case "blockquote": if (closing) { FlushCur(); } else { StartCur('p', -1, "", 0.0, ""); } break;
                        case "h1": if (closing) { FlushCur(); } else { StartCur('h', -1, "", 0.0, "font-size:20.0pt;font-weight:bold"); } break;
                        case "h2": if (closing) { FlushCur(); } else { StartCur('h', -1, "", 0.0, "font-size:16.0pt;font-weight:bold"); } break;
                        case "h3":
                        case "h4":
                        case "h5":
                        case "h6": if (closing) { FlushCur(); } else { StartCur('h', -1, "", 0.0, "font-weight:bold"); } break;
                        case "pre": if (closing) { _inPre = false; } else { FlushCur(); _inPre = true; } break;
                        case "br": break;
                        case "input":
                            if (_cur != null && attrs.Contains("type=\"checkbox\""))
                            {
                                _cur.Text += (attrs.Contains("checked") ? "\u2611 " : "\u2610 ");
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
                                Match hm = Regex.Match(attrs, "href=\"([^\"]*)\"");
                                AppendInline("<a href=\"" + hm.Groups[1].Value + "\">");
                            }
                            break;
                        default: break;
                    }
                }
                else
                {
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
                        if (_cur == null) StartCur('p', -1, "", 0.0, "");
                        _cur.Text += txt;
                    }
                }
            }
            FlushCur();

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

            foreach (string raw in lines)
            {
                string line = raw.Replace("\t", "  ");
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
                    sb.Append(EmitItem(it));
                    openOE = true;
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
                        Update(app, newXml);
                        return;
                    }
                }
            }

            int idx = pageXml.LastIndexOf("</one:Page>");
            if (idx < 0) { Error("Unexpected page XML structure."); return; }
            double estH2 = itemsCount * 22 + 20;
            string newOutline2 = "<one:Outline><one:Size width=\"620.0\" height=\"" + Num(estH2) + "\"/>" + outlineBody + "</one:Outline>";
            newXml = pageXml.Substring(0, idx) + newOutline2 + pageXml.Substring(idx);
            Update(app, newXml);
        }

        private static int CountItems(string outlineBody)
        {
            int n = 0;
            int i = outlineBody.IndexOf("<one:OE>", StringComparison.Ordinal);
            while (i >= 0) { n++; i = outlineBody.IndexOf("<one:OE>", i + 1, StringComparison.Ordinal); }
            return n;
        }

        private static void Update(OneNoteApp app, string newXml)
        {
            try
            {
                app.UpdatePageContent(newXml);
                Log("update OK");
            }
            catch (Exception ex)
            {
                Log("update exception: " + ex.Message);
                Error("OneNote update failed: " + ex.Message);
            }
        }
    }
}
