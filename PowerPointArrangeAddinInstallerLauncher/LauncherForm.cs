using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace PowerPointArrangeAddinInstallerLauncher {

    public sealed partial class LauncherForm : Form {

        private readonly string[] _languages = { "Default language", "English", "简体中文", "正體中文", "日本語" };
        private readonly int[] _languageCodes = { 0, 1033, 2052, 1028, 1041 };

        public LauncherForm() {
            InitializeComponent();

            tlpMain.AutoSize = true;
            tlpMain.Dock = DockStyle.Fill;

            AutoScaleMode = AutoScaleMode.Dpi;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            AutoSize = true;
            Font = SystemFonts.MessageBoxFont;

            foreach (var language in _languages) {
                cbxLanguage.Items.Add(language);
            }
            cbxLanguage.SelectedIndex = 0;
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            var filename = "_$_PowerPointArrangeAddinInstaller.temp";

            try {
                var stream = File.Open(filename, FileMode.CreateNew);
                File.SetAttributes(filename, File.GetAttributes(filename) | FileAttributes.Hidden);
                var w = new BinaryWriter(stream);
                w.Write(Properties.Resources.PowerPointArrangeAddinInstaller); // msi file
                w.Close();
            } catch (Exception ex) {
                ErrMsgBox($"Failed to launch installer:\r\n\r\n{ex.Message}");
                SafeDeleteFile(filename);
                Close();
                return;
            }

            var language = _languageCodes[cbxLanguage.SelectedIndex];
            var psi = new ProcessStartInfo {
                FileName = "msiexec",
                Arguments = $"/i {filename} ProductLanguage={language}"
            };

            try {
                var p = Process.Start(psi);
                Hide();
                p?.WaitForExit();
            } catch (Exception ex) {
                ErrMsgBox($"Failed to launch installer:\r\n\r\n{ex.Message}");
            } finally {
                SafeDeleteFile(filename);
                Close();
            }
        }

        private void ErrMsgBox(string text) {
            MessageBox.Show(text, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void SafeDeleteFile(string filename) {
            try {
                File.Delete(filename);
            } catch (Exception) {
                // ignored
            }
        }

        private void AdjustRegistry() {
            // {5CCB2D8E-3BC0-4B35-AB6C-B504CFAE892E}
            //  E8D2BCC5 0CB3 53B4 BAC6 5B40FCEA98E2

            // コンピューター\HKEY_CURRENT_USER\SOFTWARE\AoiHosizora\PowerPointArrangeAddin
            // コンピューター\HKEY_CLASSES_ROOT\Installer\Products\E8D2BCC50CB353B4BAC65B40FCEA98E2
            // コンピューター\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Installer\Products\E8D2BCC50CB353B4BAC65B40FCEA98E2
            // コンピューター\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\E8D2BCC50CB353B4BAC65B40FCEA98E2\InstallProperties
            // コンピューター\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{5CCB2D8E-3BC0-4B35-AB6C-B504CFAE892E}
        }

        private void BtnCancel_Click(object sender, EventArgs e) {
            Close();
        }

    }

}
