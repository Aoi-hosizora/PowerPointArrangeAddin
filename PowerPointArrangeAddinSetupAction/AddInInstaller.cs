using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Threading;

[RunInstaller(true)]
public class AddInInstaller : Installer {

    private string GetTargetDir() {
        return @"C:\Program Files\AoiHosizora\PowerPointArrangeAddin";
        var targetDir = Context.Parameters["dir"].TrimEnd('/', '\\');
        return targetDir;
    }

    private string GetVstoInstallerPath() {
        return @"C:\Program Files\AoiHosizora\PowerPointArrangeAddin\VSTOInstaller.exe";
        //return @"C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe";
        return System.IO.Path.Combine(GetTargetDir(), "VSTOInstaller.exe");
    }

    private string GetVstoFilePath() {
        return @"E:\Projects\PowerPointArrangeAddin\PowerPointArrangeAddin\bin\x64\Release\PowerPointArrangeAddin.vsto";
        return "C:\\Program Files\\AoiHosizora\\PowerPointArrangeAddin\\PowerPointArrangeAddin.vsto";
        //return "file:///C:/Program%20Files/AoiHosizora/PowerPointArrangeAddin/PowerPointArrangeAddin.vsto";
        return System.IO.Path.Combine(GetTargetDir(), "PowerPointArrangeAddin.vsto");
    }

    public override void Commit(IDictionary savedState) {
        base.Commit(savedState);
        System.Windows.Forms.MessageBox.Show(GetVstoFilePath() + ": " + (File.Exists(GetVstoFilePath()) ? "true" : "false") + "; " + GetVstoInstallerPath());
        var psi = new ProcessStartInfo {
            FileName = GetVstoInstallerPath(),
            Arguments = $"/Install \"{GetVstoFilePath()}\"",
            WorkingDirectory = GetTargetDir()
        };
        var p = Process.Start(psi); // TODO
        if (p == null) {
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

    public static void Main(string[] args) {
        var psi = new ProcessStartInfo {
            FileName = @"C:\Program Files\AoiHosizora\PowerPointArrangeAddin\VSTOInstaller.exe",
            Arguments = $"/Install \"C:\\Program Files\\AoiHosizora\\PowerPointArrangeAddin\\PowerPointArrangeAddin.vsto\"",
            WorkingDirectory = @"C:\Program Files\AoiHosizora\PowerPointArrangeAddin"
        };
        var p = Process.Start(psi);
        if (p == null) {
            throw new InstallException("Cannot execute VSTO installer.");
        }
        p.WaitForExit();
        if (p.ExitCode != 0) {
            throw new InstallException("Failed to install VSTO.");
        }
    }

}
