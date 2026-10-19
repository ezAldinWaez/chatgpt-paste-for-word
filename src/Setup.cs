using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("ChatGPT Paste for Word Setup")]
[assembly: AssemblyProduct("ChatGPT Paste for Word")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyInformationalVersion("0.1.0-alpha")]

internal sealed class SetupWindow : Form
{
    private const string AppName = "ChatGPT Paste for Word";
    private const string InstallFolder = "ChatGPTPasteForWord";
    private const string RibbonProgId = "ChatGPTPasteForWord.Ribbon";
    private const string RibbonClsid = "{E1379E36-6276-4EAB-8A0E-5A30FD4C91A7}";
    private const string RibbonAssembly = "ChatGPTPasteRibbon-0.1.0.dll";
    private const string RibbonVersion = "0.1.0.0";
    private const string PandocExeSha256 = "E83F8354C0F507222B5684797B9C5AE766F03889785995D14AAC27816EC456BA";
    private const string PandocUrl = "https://github.com/jgm/pandoc/releases/download/3.9.0.2/pandoc-3.9.0.2-windows-x86_64.zip";
    private const string PandocSha256 = "C97542F2800F446E788D9F74237856D995421AD1BB3CC8324286840C5F272D3A";
    private readonly Label status;
    private readonly Button installButton;
    private readonly Button uninstallButton;
    private readonly ProgressBar progress;

    internal SetupWindow()
    {
        Text = AppName + " Setup v0.1.0 (Alpha)";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(460, 215);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        Icon = SystemIcons.Application;

        var title = new Label { Text = AppName, Font = new Font(FontFamily.GenericSansSerif, 14, FontStyle.Bold), Left = 22, Top = 20, Width = 415, Height = 30 };
        var description = new Label { Text = "Adds a ChatGPT Paste button to Word's Home ribbon and a ChatGPT Paste tab. The keyboard shortcut also remains available.", Left = 22, Top = 57, Width = 415, Height = 45 };
        var note = new Label { Text = "First install downloads Pandoc (~41 MB); updates reuse a verified copy. Reopen Word after setup. Uninstall works offline. No admin access needed.", Left = 22, Top = 106, Width = 415, Height = 42 };
        bool installed = IsInstalled();
        status = new Label { Text = installed ? "Installed. Choose Update or Uninstall." : "Ready to install.", Left = 22, Top = 153, Width = 240, Height = 24 };
        progress = new ProgressBar { Left = 22, Top = 181, Width = 232, Height = 12, Style = ProgressBarStyle.Marquee, Visible = false };
        uninstallButton = new Button { Text = "Uninstall", Left = 266, Top = 167, Width = 82, Height = 30, Enabled = installed };
        uninstallButton.Click += UninstallClicked;
        installButton = new Button { Text = installed ? "Update" : "Install", Left = 358, Top = 167, Width = 80, Height = 30 };
        installButton.Click += InstallClicked;
        Controls.AddRange(new Control[] { title, description, note, status, progress, uninstallButton, installButton });
        FormClosing += (sender, args) => { if (progress.Visible) args.Cancel = true; };
    }

    private async void InstallClicked(object sender, EventArgs e)
    {
        if (!Environment.Is64BitOperatingSystem)
        {
            MessageBox.Show("This installer requires 64-bit Windows.", AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        installButton.Enabled = false;
        uninstallButton.Enabled = false;
        progress.Visible = true;
        try
        {
            await Task.Run(() => Install(message => BeginInvoke(new Action(() => status.Text = message))));
            progress.Visible = false;
            status.Text = "Installed. Reopen Word for the button.";
            installButton.Text = "Update";
            installButton.Enabled = true;
            uninstallButton.Enabled = true;
            int runningWord = Process.GetProcessesByName("WINWORD").Length;
            string restart = runningWord > 0
                ? "Word is still running in " + runningWord + " process(es), including possible background windows. Save your documents, close Word completely, then reopen it to load this version."
                : "Open Word to load the new ChatGPT Paste button.";
            MessageBox.Show("Installed. " + restart + " Copy a ChatGPT response, then click ChatGPT Paste on the Home ribbon or its own tab.", AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception error)
        {
            progress.Visible = false;
            status.Text = "Installation failed.";
            installButton.Enabled = true;
            uninstallButton.Enabled = IsInstalled();
            MessageBox.Show(error.Message, AppName + " Setup", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private async void UninstallClicked(object sender, EventArgs e)
    {
        if (MessageBox.Show("Remove ChatGPT Paste for Word from this Windows account?", AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        installButton.Enabled = false;
        uninstallButton.Enabled = false;
        progress.Visible = true;
        try
        {
            string result = await Task.Run(() => Uninstall(message => BeginInvoke(new Action(() => status.Text = message))));
            progress.Visible = false;
            status.Text = "Uninstalled.";
            installButton.Text = "Install";
            installButton.Enabled = true;
            MessageBox.Show(result, AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception error)
        {
            progress.Visible = false;
            status.Text = "Uninstall failed.";
            installButton.Enabled = true;
            uninstallButton.Enabled = IsInstalled();
            MessageBox.Show(error.Message, AppName + " Setup", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static bool IsInstalled()
    {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\ChatGPTPasteForWord"))
            return key != null && key.GetValue("InstallInProgress") == null;
    }

    private static void CopyResource(string resourceName, string destination)
    {
        using (Stream source = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
        {
            if (source == null) throw new InvalidOperationException("Installer resource missing: " + resourceName);
            using (Stream target = File.Create(destination)) source.CopyTo(target);
        }
    }

    private static void ExtractOne(ZipArchive archive, string entryName, string destination)
    {
        ZipArchiveEntry entry = archive.GetEntry(entryName);
        if (entry == null) throw new InvalidDataException("Pandoc download is missing " + entryName);
        using (Stream source = entry.Open())
        using (Stream target = File.Create(destination)) source.CopyTo(target);
    }

    private static string HashFile(string path)
    {
        using (SHA256 sha = SHA256.Create())
        using (Stream input = File.OpenRead(path)) return BitConverter.ToString(sha.ComputeHash(input)).Replace("-", "");
    }

    private static void StopInstalledApp(string installDir)
    {
        string expected = Path.Combine(installDir, "ChatGPTPasteHotkey.exe");
        foreach (Process process in Process.GetProcessesByName("ChatGPTPasteHotkey"))
        {
            try
            {
                if (!String.Equals(process.MainModule.FileName, expected, StringComparison.OrdinalIgnoreCase)) continue;
                process.Kill();
                process.WaitForExit(5000);
            }
            catch { }
            finally { process.Dispose(); }
        }
    }

    private static void RegisterRibbon(string installDir)
    {
        string dllPath = Path.Combine(installDir, RibbonAssembly);
        string assembly = AssemblyName.GetAssemblyName(dllPath).FullName;
        string codeBase = new Uri(dllPath).AbsoluteUri;
        foreach (RegistryView view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
            {
                using (RegistryKey key = root.CreateSubKey(@"Software\Classes\" + RibbonProgId)) key.SetValue("", "ChatGPT Paste for Word Ribbon");
                using (RegistryKey key = root.CreateSubKey(@"Software\Classes\" + RibbonProgId + @"\CLSID")) key.SetValue("", RibbonClsid);
                string clsidRoot = @"Software\Classes\CLSID\" + RibbonClsid;
                using (RegistryKey key = root.CreateSubKey(clsidRoot)) key.SetValue("", "ChatGPT Paste for Word Ribbon");
                using (RegistryKey key = root.CreateSubKey(clsidRoot + @"\InprocServer32"))
                {
                    key.SetValue("", "mscoree.dll");
                    key.SetValue("ThreadingModel", "Both");
                    key.SetValue("Class", "ChatGPTPasteRibbon");
                    key.SetValue("Assembly", assembly);
                    key.SetValue("RuntimeVersion", "v4.0.30319");
                    key.SetValue("CodeBase", codeBase);
                }
                using (RegistryKey key = root.CreateSubKey(clsidRoot + @"\InprocServer32\" + RibbonVersion))
                {
                    key.SetValue("Class", "ChatGPTPasteRibbon");
                    key.SetValue("Assembly", assembly);
                    key.SetValue("RuntimeVersion", "v4.0.30319");
                    key.SetValue("CodeBase", codeBase);
                }
                using (RegistryKey key = root.CreateSubKey(clsidRoot + @"\ProgId")) key.SetValue("", RibbonProgId);
                using (RegistryKey key = root.CreateSubKey(clsidRoot + @"\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}")) { }
                RegisterRibbonRecord(root, "{289E9AF1-4973-11D1-AE81-00A0C90F26F4}", "Extensibility.ext_ConnectMode", assembly, codeBase);
                RegisterRibbonRecord(root, "{289E9AF2-4973-11D1-AE81-00A0C90F26F4}", "Extensibility.ext_DisconnectMode", assembly, codeBase);
                using (RegistryKey key = root.CreateSubKey(@"Software\Microsoft\Office\Word\Addins\" + RibbonProgId))
                {
                    key.SetValue("FriendlyName", AppName);
                    key.SetValue("Description", "Paste ChatGPT content with editable equations and formatting.");
                    key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
                }
            }
        }
    }

    private static void RegisterRibbonRecord(RegistryKey root, string recordId, string className, string assembly, string codeBase)
    {
        using (RegistryKey key = root.CreateSubKey(@"Software\Classes\Record\" + recordId + @"\" + RibbonVersion))
        {
            key.SetValue("Class", className);
            key.SetValue("Assembly", assembly);
            key.SetValue("RuntimeVersion", "v4.0.30319");
            key.SetValue("CodeBase", codeBase);
        }
    }

    private static void UnregisterRibbon()
    {
        foreach (RegistryView view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
            {
                root.DeleteSubKeyTree(@"Software\Microsoft\Office\Word\Addins\" + RibbonProgId, false);
                root.DeleteSubKeyTree(@"Software\Classes\" + RibbonProgId, false);
                root.DeleteSubKeyTree(@"Software\Classes\CLSID\" + RibbonClsid, false);
                foreach (string recordId in new[] { "{289E9AF1-4973-11D1-AE81-00A0C90F26F4}", "{289E9AF2-4973-11D1-AE81-00A0C90F26F4}" })
                {
                    string recordRoot = @"Software\Classes\Record\" + recordId;
                    using (RegistryKey versions = root.OpenSubKey(recordRoot))
                    {
                        if (versions == null) continue;
                        foreach (string version in versions.GetSubKeyNames())
                        {
                            string path = recordRoot + @"\" + version;
                            using (RegistryKey key = root.OpenSubKey(path))
                                if (key == null || !Convert.ToString(key.GetValue("Assembly", "")).StartsWith("ChatGPTPasteRibbon-", StringComparison.Ordinal)) continue;
                            root.DeleteSubKeyTree(path, false);
                        }
                    }
                }
            }
        }
    }

    private static string Uninstall(Action<string> report)
    {
        report("Removing ChatGPT Paste...");
        string script = Path.Combine(Path.GetTempPath(), "ChatGPTPasteForWord-Uninstall-" + Guid.NewGuid().ToString("N") + ".ps1");
        string cleanup = Path.Combine(Path.GetTempPath(), "ChatGPTPasteForWord-Cleanup-" + Guid.NewGuid().ToString("N") + ".ps1");
        bool cleanupStarted = false;
        CopyResource("Uninstall.ps1", script);
        CopyResource("Cleanup-Word-Paste.ps1", cleanup);
        try
        {
            string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
            var start = new ProcessStartInfo(powershell, "-NoProfile -ExecutionPolicy Bypass -File \"" + script + "\" -CleanupSource \"" + cleanup + "\" -SkipCleanupLaunch")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            using (Process process = Process.Start(start))
            {
                if (!process.WaitForExit(60000)) { process.Kill(); throw new TimeoutException("Uninstall timed out. Run setup again to finish removing it."); }
                if (process.ExitCode != 0) throw new InvalidOperationException("Uninstall could not finish. Close Word, then run setup again.");
                string installDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), InstallFolder);
                if (!Directory.Exists(installDir)) return "ChatGPT Paste for Word was removed completely.";
                var cleanupStart = new ProcessStartInfo(powershell, "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"" + cleanup + "\" -InstallDir \"" + installDir + "\"")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                using (Process helper = Process.Start(cleanupStart))
                {
                    if (helper == null) throw new InvalidOperationException("The add-in was disabled, but file cleanup could not start.");
                    cleanupStarted = true;
                }
                return "ChatGPT Paste was removed from Word and startup. Close every Word window to finish deleting files currently in use.";
            }
        }
        finally { try { File.Delete(script); } catch { } if (!cleanupStarted) try { File.Delete(cleanup); } catch { } }
    }

    private static void Install(Action<string> report)
    {
        string localData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string installDir = Path.Combine(localData, InstallFolder);
        string staging = Path.Combine(localData, InstallFolder + ".stage." + Guid.NewGuid().ToString("N"));
        bool updating = Directory.Exists(installDir);
        bool hadUninstallEntry;
        using (RegistryKey existing = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\ChatGPTPasteForWord")) hadUninstallEntry = existing != null;
        bool movedNew = false;
        try
        {
            using (RegistryKey pending = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\ChatGPTPasteForWord"))
                pending.SetValue("InstallInProgress", 1, RegistryValueKind.DWord);
            Directory.CreateDirectory(staging);
            string existingPandoc = Path.Combine(installDir, "pandoc.exe");
            if (File.Exists(existingPandoc) && String.Equals(HashFile(existingPandoc), PandocExeSha256, StringComparison.OrdinalIgnoreCase))
            {
                report("Reusing verified Pandoc...");
                File.Copy(existingPandoc, Path.Combine(staging, "pandoc.exe"));
                File.Copy(Path.Combine(installDir, "PANDOC-COPYING.rtf"), Path.Combine(staging, "PANDOC-COPYING.rtf"));
                File.Copy(Path.Combine(installDir, "PANDOC-COPYRIGHT.txt"), Path.Combine(staging, "PANDOC-COPYRIGHT.txt"));
            }
            else
            {
                report("Downloading Pandoc...");
                string zipPath = Path.Combine(staging, "pandoc.zip");
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                using (var client = new WebClient()) client.DownloadFile(PandocUrl, zipPath);
                report("Verifying the download...");
                if (!String.Equals(HashFile(zipPath), PandocSha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Pandoc download failed its SHA-256 check. Installation stopped.");
                using (ZipArchive archive = ZipFile.OpenRead(zipPath))
                {
                    const string prefix = "pandoc-3.9.0.2/";
                    ExtractOne(archive, prefix + "pandoc.exe", Path.Combine(staging, "pandoc.exe"));
                    ExtractOne(archive, prefix + "COPYING.rtf", Path.Combine(staging, "PANDOC-COPYING.rtf"));
                    ExtractOne(archive, prefix + "COPYRIGHT.txt", Path.Combine(staging, "PANDOC-COPYRIGHT.txt"));
                }
                File.Delete(zipPath);
            }
            CopyResource("ChatGPTPasteHotkey.exe", Path.Combine(staging, "ChatGPTPasteHotkey.exe"));
            CopyResource(RibbonAssembly, Path.Combine(staging, RibbonAssembly));
            CopyResource("Prepare-Word-Paste.ps1", Path.Combine(staging, "Prepare-Word-Paste.ps1"));
            CopyResource("Uninstall.ps1", Path.Combine(staging, "Uninstall.ps1"));
            CopyResource("Cleanup-Word-Paste.ps1", Path.Combine(staging, "Cleanup-Word-Paste.ps1"));

            report("Installing for this Windows account...");
            StopInstalledApp(installDir);
            if (updating)
            {
                // An open Word window can hold the previous ribbon DLL open.
                // Copy the new version beside it instead of moving its folder.
                foreach (string file in Directory.GetFiles(staging))
                    File.Copy(file, Path.Combine(installDir, Path.GetFileName(file)), true);
            }
            else
            {
                Directory.Move(staging, installDir);
                movedNew = true;
            }

            RegisterRibbon(installDir);

            foreach (string oldRibbon in Directory.GetFiles(installDir, "ChatGPTPasteRibbon-*.dll"))
                if (!String.Equals(Path.GetFileName(oldRibbon), RibbonAssembly, StringComparison.OrdinalIgnoreCase))
                    try { File.Delete(oldRibbon); } catch (IOException) { } catch (UnauthorizedAccessException) { }

            using (RegistryKey run = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run"))
                run.SetValue("ChatGPTPasteForWord", "\"" + Path.Combine(installDir, "ChatGPTPasteHotkey.exe") + "\"");
            using (RegistryKey uninstall = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\ChatGPTPasteForWord"))
            {
                uninstall.SetValue("DisplayName", AppName + " (unofficial)");
                uninstall.SetValue("DisplayVersion", "0.1.0");
                uninstall.SetValue("Publisher", "Independent utility");
                uninstall.SetValue("InstallLocation", installDir);
                string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
                uninstall.SetValue("UninstallString", "\"" + powershell + "\" -NoProfile -ExecutionPolicy Bypass -File \"" + Path.Combine(installDir, "Uninstall.ps1") + "\"");
                uninstall.SetValue("NoModify", 1, RegistryValueKind.DWord);
                uninstall.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                uninstall.DeleteValue("InstallInProgress", false);
            }
            report("Starting ChatGPT Paste...");
            var start = new ProcessStartInfo(Path.Combine(installDir, "ChatGPTPasteHotkey.exe")) { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden };
            using (Process app = Process.Start(start))
            {
                Thread.Sleep(1000);
                if (app.HasExited) throw new InvalidOperationException("The paste shortcut could not start. Another app may be using Ctrl+Alt+Shift+V.");
            }
        }
        catch
        {
            if (!updating)
            {
                UnregisterRibbon();
                StopInstalledApp(installDir);
                if (movedNew && Directory.Exists(installDir)) Directory.Delete(installDir, true);
                using (RegistryKey run = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                    if (run != null) run.DeleteValue("ChatGPTPasteForWord", false);
                Registry.CurrentUser.DeleteSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\ChatGPTPasteForWord", false);
            }
            else if (File.Exists(Path.Combine(installDir, RibbonAssembly)))
            {
                // Keep the installed add-in registered even if a later step fails.
                try { RegisterRibbon(installDir); } catch { }
            }
            if (!hadUninstallEntry)
                Registry.CurrentUser.DeleteSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\ChatGPTPasteForWord", false);
            else
                using (RegistryKey uninstall = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\ChatGPTPasteForWord", true))
                    if (uninstall != null) uninstall.DeleteValue("InstallInProgress", false);
            throw;
        }
        finally { if (Directory.Exists(staging)) Directory.Delete(staging, true); }
    }

    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length == 1 && String.Equals(args[0], "/uninstall", StringComparison.OrdinalIgnoreCase))
        {
            try { Uninstall(message => { }); Environment.ExitCode = 0; }
            catch (Exception error)
            {
                File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ChatGPTPasteForWord-uninstall-error.txt"), error.ToString());
                Environment.ExitCode = 1;
            }
            return;
        }
        if (args.Length == 1 && String.Equals(args[0], "/quiet", StringComparison.OrdinalIgnoreCase))
        {
            try { Install(message => { }); Environment.ExitCode = 0; }
            catch (Exception error)
            {
                File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ChatGPTPasteForWord-setup-error.txt"), error.ToString());
                Environment.ExitCode = 1;
            }
            return;
        }
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new SetupWindow());
    }
}
