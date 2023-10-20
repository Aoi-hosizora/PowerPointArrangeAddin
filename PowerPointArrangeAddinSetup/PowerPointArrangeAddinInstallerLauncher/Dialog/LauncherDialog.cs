using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

#nullable enable

namespace PowerPointArrangeAddinInstallerLauncher.Dialog {

    public sealed partial class LauncherDialog : Form {

        private readonly string[] _languages = { "English", "简体中文", "正體中文", "日本語" };
        private readonly int[] _languageCodes = { 1033, 2052, 1028, 1041 };
        private readonly string[] _args;

        public LauncherDialog(string[]? args) {
            args ??= new string[] { };
            _args = args.Length == 0 || args.Contains("-") ? args : args.Append("-").ToArray();

            InitializeComponent();

            tlpMain.AutoSize = true;
            tlpMain.Dock = DockStyle.Fill;

            AutoScaleMode = AutoScaleMode.Dpi;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            AutoSize = true;
            Font = SystemFonts.MessageBoxFont;

            if (_args.Length > 0) {
                lblHint.Text += $"\r\n\"msiexec {string.Join(" ", _args)}\"";
            }

            foreach (var language in _languages) {
                cboLanguage.Items.Add(language);
            }
            cboLanguage.SelectedIndex = GetCurrentLanguage();
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            var filename = "_$_PowerPointArrangeAddinInstaller.tmp";

            try {
                var stream = File.Open(filename, FileMode.CreateNew);
                File.SetAttributes(filename, File.GetAttributes(filename) | FileAttributes.Hidden);
                var w = new BinaryWriter(stream);
                w.Write(Properties.Resources.PowerPointArrangeAddinInstaller); // msi file
                w.Close();
            } catch (Exception ex) {
                ErrMsgBox($"Failed to launch installer:\r\n\r\n{ex.Message}", Text);
                SafeDeleteFile(filename);
                Close();
                return;
            }

            var language = _languageCodes[cboLanguage.SelectedIndex];
            var arguments = new List<string>();
            if (_args.Length == 0) {
                arguments.Add($"/i {filename}");
            } else {
                arguments.AddRange(_args.Select(arg => arg != "-" ? arg : filename));
            }
            arguments.Add($"ProductLanguage={language}");

            try {
                var psi = new ProcessStartInfo { FileName = "msiexec", Arguments = string.Join(" ", arguments) };
                var p = Process.Start(psi);
                Hide();
                p?.WaitForExit();
            } catch (Exception ex) {
                ErrMsgBox($"Failed to launch installer:\r\n\r\n{ex.Message}", Text);
            } finally {
                SafeDeleteFile(filename);
                Close();
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e) {
            Close();
        }

        private static void ErrMsgBox(string text, string title) {
            MessageBox.Show(text, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void SafeDeleteFile(string filename) {
            try {
                File.Delete(filename);
            } catch (Exception) {
                // ignored
            }
        }

        private static int GetCurrentLanguage() {
            var name = Thread.CurrentThread.CurrentCulture.Name.ToLower();
            return name switch {
                "zh" or "zh-hans" or "zh-chs" or "zh-cn" or "zh-sg" => 1,
                "zh-hant" or "zh-cht" or "zh-tw" or "zh-hk" or "zh-mo" => 2,
                "ja" or "ja-jp" => 3,
                _ => 0
            };
        }
    }

}
