using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

[RunInstaller(true)]
public class AddInInstaller : Installer {

    private string GetTargetDir() {
        //return @"C:\Program Files\AoiHosizora\PowerPointArrangeAddin";
        var targetDir = Context.Parameters["dir"].TrimEnd('/', '\\');
        return targetDir;
    }

    private string GetVstoInstallerPath() {
        //return @"C:\Program Files\AoiHosizora\PowerPointArrangeAddin\VSTOInstaller.exe";
        //return @"C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe";
        return System.IO.Path.Combine(GetTargetDir(), "VSTOInstaller.exe");
    }

    private string GetVstoFilePath() {
        //return @"E:\Projects\PowerPointArrangeAddin\PowerPointArrangeAddin\bin\x64\Release\PowerPointArrangeAddin.vsto";
        //return "C:\\Program Files\\AoiHosizora\\PowerPointArrangeAddin\\PowerPointArrangeAddin.vsto";
        //return "file:///C:/Program%20Files/AoiHosizora/PowerPointArrangeAddin/PowerPointArrangeAddin.vsto";
        return System.IO.Path.Combine(GetTargetDir(), "PowerPointArrangeAddin.vsto");
    }

    public override void Install(IDictionary savedState) {
        base.Install(savedState);
        //Directory.SetCurrentDirectory(GetTargetDir());
        //MessageBox.Show(Path.GetDirectoryName
        //    (Assembly.GetExecutingAssembly().Location));
        //MessageBox.Show(GetVstoFilePath() + ": " + (File.Exists(GetVstoFilePath()) ? "true" : "false") + "; " + GetVstoInstallerPath());
        //var psi = new ProcessStartInfo {
        //    FileName = GetVstoInstallerPath(),
        //    Arguments = $"/Install \"{GetVstoFilePath()}\"",
        //    WorkingDirectory = GetTargetDir()
        //}
        var psi = new ProcessStartInfo {
            FileName = @"C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
            Arguments = @"/Install ""C:\Program Files\AoiHosizora\PowerPointArrangeAddin\PowerPointArrangeAddin.vsto""",
            WorkingDirectory = @"C:\Program Files\AoiHosizora\PowerPointArrangeAddin",
            // EnvironmentVariables = { { "APPDATA", "" } },
            // Environment = { { "PATH", $@"C:\Program Files\AoiHosizora\PowerPointArrangeAddin; {Environment.GetEnvironmentVariable("PATH")}" } },
            UseShellExecute = false,
            Verb = "runas"
        };
        var p = new Process { StartInfo = psi };
        var ok = p.Start(); // TODO
        if (!ok) {
            throw new InstallException("Cannot execute VSTO installer.");
        }
        p.WaitForExit();
        if (p.ExitCode != 0) {
            throw new InstallException("Failed to install VSTO.");
        }
    }

    public override void Uninstall(IDictionary savedState) {
        base.Uninstall(savedState);
        // Thread.Sleep(2000);
        // var psi = new ProcessStartInfo {
        //     FileName = GetVstoInstallerPath(),
        //     Arguments = $"/Uninstall {GetVstoFilePath()}",
        //     WorkingDirectory = GetTargetDir()
        // };
        // var p = Process.Start(psi);
        // if (p == null) {
        //     throw new InstallException("Cannot execute VSTO uninstaller.");
        // }
        // p?.WaitForExit();
        // if (p.ExitCode != 0) {
        //     throw new InstallException("Failed to uninstall VSTO.");
        // }
    }

    //public void Install() {
    //    System.Windows.Forms.MessageBox.Show(GetVstoFilePath() + ": " + (File.Exists(GetVstoFilePath()) ? "true" : "false") + "; " + GetVstoInstallerPath());
    //    var psi = new ProcessStartInfo {
    //        FileName = GetVstoInstallerPath(),
    //        Arguments = $"/Install \"{GetVstoFilePath()}\"",
    //        WorkingDirectory = GetTargetDir()
    //    };
    //    var p = Process.Start(psi); // TODO
    //    if (p == null) {
    //        throw new InstallException("Cannot execute VSTO installer.");
    //    }
    //    p.WaitForExit();
    //    if (p.ExitCode != 0) {
    //        throw new InstallException("Failed to install VSTO.");
    //    }
    //}

    //public void Uninstall() {
    //    System.Windows.Forms.MessageBox.Show(GetVstoFilePath() + ": " + (File.Exists(GetVstoFilePath()) ? "true" : "false") + "; " + GetVstoInstallerPath());
    //    var psi = new ProcessStartInfo {
    //        FileName = GetVstoInstallerPath(),
    //        Arguments = $"/Uninstall \"{GetVstoFilePath()}\"",
    //        WorkingDirectory = GetTargetDir()
    //    };
    //    var p = Process.Start(psi); // TODO
    //    if (p == null) {
    //        throw new InstallException("Cannot execute VSTO installer.");
    //    }
    //    p.WaitForExit();
    //    if (p.ExitCode != 0) {
    //        throw new InstallException("Failed to install VSTO.");
    //    }
    //}

    //public static void Main(string[] args) {
    //    var installer = new AddInInstaller();
    //    if (args.Length >= 2 && (args[1].ToLower() == "/u" || args[1].ToLower() == "/uninstaller")) {
    //        installer.Uninstall();
    //    } else {
    //        installer.Install();
    //    }
    //}

}
