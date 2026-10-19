using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;

[assembly: System.Reflection.AssemblyTitle("ChatGPT Paste for Word Ribbon")]
[assembly: System.Reflection.AssemblyProduct("ChatGPT Paste for Word")]
[assembly: System.Reflection.AssemblyVersion("1.5.9.0")]

[ComVisible(true)]
[Guid("E1379E36-6276-4EAB-8A0E-5A30FD4C91A7")]
[ProgId("ChatGPTPasteForWord.Ribbon")]
[ClassInterface(ClassInterfaceType.AutoDual)]
public sealed class ChatGPTPasteRibbon : IDTExtensibility2, IRibbonExtensibility
{
    private object wordApplication;
    private Control dispatcher;

    public string Version { get { return "1.5.9"; } }

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        wordApplication = application;
        dispatcher = new Control();
        IntPtr ignored = dispatcher.Handle;
        try { ((COMAddIn)addInInst).Object = this; }
        catch (Exception error) { Log("Could not expose shortcut bridge: " + error.Message); }
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) { wordApplication = null; if (dispatcher != null) dispatcher.Dispose(); dispatcher = null; }
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { wordApplication = null; }

    public string GetCustomUI(string ribbonId)
    {
        return "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
               "<ribbon><tabs>" +
               "<tab idMso=\"TabHome\"><group id=\"ChatGPTPasteHomeGroup\" label=\"ChatGPT\">" +
               "<button id=\"ChatGPTPasteHomeButton\" label=\"ChatGPT Paste\" size=\"large\" imageMso=\"Paste\" onAction=\"OnPasteClick\" screentip=\"Paste from ChatGPT\" supertip=\"Keep headings, lists, tables, editable equations, and the cursor's font and direction. Version 1.5.9.\"/>" +
               "</group></tab>" +
               "<tab id=\"ChatGPTPasteTab\" label=\"ChatGPT Paste\"><group id=\"ChatGPTPasteTabGroup\" label=\"Paste\">" +
               "<button id=\"ChatGPTPasteTabButton\" label=\"Paste from ChatGPT\" size=\"large\" imageMso=\"Paste\" onAction=\"OnPasteClick\"/>" +
               "</group></tab>" +
               "</tabs></ribbon></customUI>";
    }

    public void OnPasteClick(IRibbonControl control)
    {
        Log("Button clicked");
        PasteFromChatGPT();
    }

    // Called by the tray shortcut. Marshal the work onto the Word UI thread
    // through the control created when the add-in connected.
    public void QueuePaste()
    {
        if (dispatcher == null || dispatcher.IsDisposed) throw new InvalidOperationException("Word paste dispatcher is unavailable.");
        dispatcher.BeginInvoke(new Action(delegate
        {
            Log("Shortcut paste started");
            PasteFromChatGPT();
        }));
    }

    public void PasteFromChatGPT()
    {
        try
        {
            string helper = Path.Combine(Path.GetDirectoryName(GetType().Assembly.Location), "Prepare-Word-Paste.ps1");
            if (!File.Exists(helper)) throw new FileNotFoundException("Clipboard converter is missing. Reinstall ChatGPT Paste for Word.");
            var start = new ProcessStartInfo
            {
                FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
                Arguments = "-NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File \"" + helper + "\" -Quiet",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true
            };
            string error;
            int exitCode;
            using (Process converter = Process.Start(start))
            {
                error = converter.StandardError.ReadToEnd();
                if (!converter.WaitForExit(15000))
                {
                    converter.Kill();
                    throw new TimeoutException("Equation conversion timed out. The Word document was not changed.");
                }
                exitCode = converter.ExitCode;
            }
            if (exitCode != 0 && exitCode != 2) throw new InvalidOperationException(error.Trim());
            if (wordApplication == null) throw new InvalidOperationException("Word is not connected to the ChatGPT Paste add-in.");
            dynamic application = wordApplication;
            dynamic selection = application.Selection;
            dynamic document = application.ActiveDocument;
            WordPasteContext context = WordPasteFormatting.Capture(selection, document);
            if (Clipboard.ContainsData(DataFormats.Html))
            {
                object missing = Type.Missing;
                selection.PasteSpecial(missing, missing, missing, missing, 10, missing, missing); // wdPasteHTML
            }
            else selection.Paste();
            int adjusted = WordPasteFormatting.Normalize(document, selection, context);
            Log("Paste completed; normalized " + adjusted + " paragraphs; direction=" + context.ReadingOrder);
        }
        catch (Exception error)
        {
            Log("ERROR " + error);
            MessageBox.Show(error.Message, "ChatGPT Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static void Log(string message)
    {
        try
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ChatGPTPasteForWord");
            Directory.CreateDirectory(folder);
            File.AppendAllText(Path.Combine(folder, "ribbon.log"), DateTime.Now.ToString("s") + " " + message + Environment.NewLine);
        }
        catch { }
    }
}
