using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Reflection;
using System.Text;

[assembly: AssemblyTitle("ChatGPT Paste for Word")]
[assembly: AssemblyProduct("ChatGPT Paste for Word")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyInformationalVersion("0.1.0-alpha")]

internal static class NativeMethods
{
    internal delegate bool EnumChildProc(IntPtr hwnd, IntPtr parameter);
    [DllImport("user32.dll")] internal static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint virtualKey);
    [DllImport("user32.dll")] internal static extern bool UnregisterHotKey(IntPtr window, int id);
    [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll")] internal static extern bool EnumChildWindows(IntPtr parent, EnumChildProc callback, IntPtr parameter);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetClassName(IntPtr window, StringBuilder name, int length);
    [DllImport("oleacc.dll")] internal static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid interfaceId, [MarshalAs(UnmanagedType.Interface)] out object result);

    internal static object ForegroundWordWindow()
    {
        object found = null;
        EnumChildWindows(GetForegroundWindow(), delegate(IntPtr child, IntPtr parameter)
        {
            var name = new StringBuilder(64);
            GetClassName(child, name, name.Capacity);
            if (name.ToString() != "_WwG") return true;
            Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
            object candidate;
            if (AccessibleObjectFromWindow(child, 0xFFFFFFF0, ref dispatch, out candidate) == 0 && candidate != null)
            {
                found = candidate;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}

internal sealed class ChatGPTPasteWindow : Form
{
    private const int HotkeyMessage = 0x0312;
    private const int HotkeyId = 0x4721;
    private const uint ControlAltShift = 0x0002 | 0x0001 | 0x0004;
    private readonly NotifyIcon trayIcon;
    private bool pasting;

    internal ChatGPTPasteWindow()
    {
        ShowInTaskbar = false;
        WindowState = FormWindowState.Minimized;
        var menu = new ContextMenuStrip();
        menu.Items.Add("ChatGPT Paste: Ctrl+Alt+Shift+V");
        menu.Items.Add("Exit", null, (sender, args) => Close());
        trayIcon = new NotifyIcon { Icon = SystemIcons.Application, Text = "ChatGPT Paste v0.1.0 (Alpha)", Visible = true, ContextMenuStrip = menu };
        if (!NativeMethods.RegisterHotKey(Handle, HotkeyId, ControlAltShift, 0x56))
        {
            trayIcon.Dispose();
            throw new InvalidOperationException("Ctrl+Alt+Shift+V is already in use.");
        }
    }

    protected override void SetVisibleCore(bool value) { base.SetVisibleCore(false); }

    protected override void WndProc(ref Message message)
    {
        if (message.Msg == HotkeyMessage && message.WParam.ToInt32() == HotkeyId)
        {
            Log("Hotkey received");
            BeginInvoke(new Action(PasteFromChatGPT));
            return;
        }
        base.WndProc(ref message);
    }

    private static uint ForegroundWordPid()
    {
        uint processId;
        NativeMethods.GetWindowThreadProcessId(NativeMethods.GetForegroundWindow(), out processId);
        if (processId == 0) return 0;
        try { return String.Equals(Process.GetProcessById((int)processId).ProcessName, "WINWORD", StringComparison.OrdinalIgnoreCase) ? processId : 0; }
        catch { return 0; }
    }

    private static bool PasteHtmlInWord(uint wordPid)
    {
        object window = NativeMethods.ForegroundWordWindow();
        if (window == null) { Log("No foreground Word document object"); return false; }
        try
        {
            dynamic activeWindow = window;
            uint activePid;
            NativeMethods.GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(activeWindow.Hwnd)), out activePid);
            if (activePid != wordPid) { Log("Word window PID changed"); return false; }
            object selection = activeWindow.Selection;
            try
            {
                dynamic target = selection;
                dynamic document = activeWindow.Document;
                WordPasteContext context = WordPasteFormatting.Capture(target, document);
                object missing = Type.Missing;
                target.PasteSpecial(missing, missing, missing, missing, 10, missing, missing); // wdPasteHTML
                int adjusted = WordPasteFormatting.Normalize(document, target, context);
                Log("Paste completed; normalized " + adjusted + " paragraphs; direction=" + context.ReadingOrder);
                return true;
            }
            finally { if (selection != null) Marshal.ReleaseComObject(selection); }
        }
        finally { Marshal.ReleaseComObject(window); }
    }

    private static bool PasteViaRibbon(uint wordPid)
    {
        object window = NativeMethods.ForegroundWordWindow();
        if (window == null) return false;
        try
        {
            dynamic activeWindow = window;
            uint activePid;
            NativeMethods.GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(activeWindow.Hwnd)), out activePid);
            if (activePid != wordPid) return false;
            dynamic addIn;
            try { addIn = activeWindow.Application.COMAddIns.Item("ChatGPTPasteForWord.Ribbon"); }
            catch { return false; }
            if (!Convert.ToBoolean(addIn.Connect)) throw new InvalidOperationException("Word is not using the current ChatGPT Paste add-in. Close every Word window, then reopen Word.");
            dynamic bridge = addIn.Object;
            if (bridge == null) throw new InvalidOperationException("This Word window is still using an old ChatGPT Paste add-in. Close every Word window, then reopen Word.");
            string version;
            try { version = Convert.ToString(bridge.Version); }
            catch { throw new InvalidOperationException("This Word window is still using an old ChatGPT Paste add-in. Close every Word window, then reopen Word."); }
            if (version != "0.1.0") throw new InvalidOperationException("This Word window has ChatGPT Paste " + version + ". Close every Word window, then reopen Word to load version 0.1.0.");
            Log("Using Word ribbon bridge");
            bridge.QueuePaste();
            return true;
        }
        finally { Marshal.ReleaseComObject(window); }
    }

    private void PasteFromChatGPT()
    {
        if (pasting) return;
        uint wordPid = ForegroundWordPid();
        if (wordPid == 0) { Log("No foreground Word process"); return; }
        Log("Target Word PID=" + wordPid);
        pasting = true;
        try
        {
            if (PasteViaRibbon(wordPid)) return;
            string helper = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Prepare-Word-Paste.ps1");
            if (!File.Exists(helper)) throw new FileNotFoundException("Clipboard helper was not found.", helper);
            var start = new ProcessStartInfo();
            start.FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
            start.Arguments = "-NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File \"" + helper + "\" -Quiet";
            start.UseShellExecute = false;
            start.CreateNoWindow = true;
            start.RedirectStandardError = true;
            string errorText;
            int exitCode;
            using (Process converter = Process.Start(start))
            {
                errorText = converter.StandardError.ReadToEnd();
                if (!converter.WaitForExit(15000))
                {
                    converter.Kill();
                    throw new TimeoutException("Equation conversion timed out. The Word document was not changed.");
                }
                exitCode = converter.ExitCode;
            }
            if (exitCode != 0 && exitCode != 2)
                throw new InvalidOperationException(errorText.Trim());
            if (ForegroundWordPid() != wordPid)
                throw new InvalidOperationException("Word lost focus during conversion. Click in your document and press the shortcut again.");
            if (!Clipboard.ContainsData(DataFormats.Html) || !PasteHtmlInWord(wordPid))
            {
                Log("Using ordinary Word paste fallback");
                SendKeys.SendWait("^v");
            }
        }
        catch (Exception error)
        {
            Log("ERROR " + error);
            MessageBox.Show(error.Message, "ChatGPT Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally { pasting = false; }
    }

    private static void Log(string message)
    {
        try
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ChatGPTPasteForWord");
            Directory.CreateDirectory(folder);
            File.AppendAllText(Path.Combine(folder, "hotkey.log"), DateTime.Now.ToString("s") + " " + message + Environment.NewLine);
        }
        catch { }
    }

    protected override void OnFormClosed(FormClosedEventArgs args)
    {
        NativeMethods.UnregisterHotKey(Handle, HotkeyId);
        trayIcon.Visible = false;
        trayIcon.Dispose();
        base.OnFormClosed(args);
    }
}

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        bool isFirst;
        using (var lockName = new Mutex(true, "Local\\ChatGPTPasteHotkey", out isFirst))
        {
            if (!isFirst) return;
            Application.EnableVisualStyles();
            try { Application.Run(new ChatGPTPasteWindow()); }
            catch (Exception error) { MessageBox.Show(error.Message, "ChatGPT Paste", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
    }
}
