using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Win32;

#nullable enable

namespace PowerPointArrangeAddinInstallerAction.Helper {

    internal class InstallerHelper {

        public InstallerHelper(string? productCode, string? originalDatabase, string? installFolder) {
            _productCode = productCode ?? throw new NullReferenceException(nameof(productCode));
            _originalDatabase = originalDatabase ?? throw new NullReferenceException(nameof(originalDatabase));
            _installFolder = installFolder ?? throw new NullReferenceException(nameof(installFolder));
        }

        private readonly string _productCode;
        private readonly string _originalDatabase;
        private readonly string _installFolder;

        #region Public Methods

        public void CopyInstaller() {
            InstallInformation information;
            try {
                information = GetInstallInformation();
            } catch (Exception ex) {
                throw new Exception($"Failed to get installer information:\r\n\r\n{ex.Message}");
            }

            try {
                CopyInstallerToFolder(information);
            } catch (Exception ex) {
                throw new Exception($"Failed to copy installer:\r\n\r\n{ex.Message}");
            }

            try {
                AdjustRegistryForInstaller(information);
            } catch (Exception ex) {
                throw new Exception($"Failed to update registry:\r\n\r\n{ex.Message}");
            }
        }

        public void DeleteInstaller() {
            InstallInformation information;
            try {
                information = GetInstallInformation();
            } catch (Exception ex) {
                throw new Exception($"Failed to get installer information:\r\n\r\n{ex.Message}");
            }

            try {
                DeleteInstallerFromFolder(information);
            } catch (Exception ex) {
                throw new Exception($"Failed to delete installer:\r\n\r\n{ex.Message}");
            }
        }

        #endregion

        #region Helper Methods For Installer

        private struct InstallInformation {
            public string ProductCode { get; set; }
            public string RegistryCode { get; set; }
            public string CurrentFolder { get; set; }
            public string InstallFolder { get; set; }
        }

        private InstallInformation GetInstallInformation() {
            var productCode = _productCode.Trim();
            if (productCode.Length != 38) {
                throw new Exception("the product code has wrong format");
            }
            static string ReverseString(string s, IEnumerable<(int, int)> indices) =>
                indices.Aggregate("", (curr, u) => curr + new string(s.Substring(u.Item1, u.Item2).Reverse().ToArray()));
            var registryCode = ReverseString(productCode, // {B83AEC80-3D6D-44D7-8BAA-9C23B6DC2066} => 08CEA38BD6D37D44B8AAC9326BCD0266
                new[] { (1, 8), (10, 4), (15, 4), (20, 2), (22, 2), (25, 2), (27, 2), (29, 2), (31, 2), (33, 2), (35, 2) });

            var currentFolder = Path.GetDirectoryName(_originalDatabase);
            if (currentFolder == null) {
                throw new Exception("the original installer is not found");
            }
            if (!currentFolder.EndsWith("\\")) {
                currentFolder += "\\";
            }

            var installFolder = _installFolder;
            if (!installFolder.EndsWith("\\")) {
                installFolder += "\\";
            }

            return new InstallInformation {
                ProductCode = productCode,
                RegistryCode = registryCode,
                CurrentFolder = currentFolder,
                InstallFolder = installFolder,
            };
        }

        private const string OldInstallerFilename = "_$_PowerPointArrangeAddinInstaller.tmp"; // <<<
        private const string NewInstallerFilename = "PowerPointArrangeAddinInstaller.msi";

        private void CopyInstallerToFolder(InstallInformation information) {
            var oldInstallerPath = Path.Combine(information.CurrentFolder, OldInstallerFilename);
            var newInstallerPath = Path.Combine(information.InstallFolder, NewInstallerFilename);
            File.Copy(oldInstallerPath, newInstallerPath, true);
            var installerFileAttribute = File.GetAttributes(newInstallerPath) & ~FileAttributes.Hidden;
            File.SetAttributes(newInstallerPath, installerFileAttribute);
        }

        private void DeleteInstallerFromFolder(InstallInformation information) {
            var newInstallerPath = Path.Combine(information.InstallFolder, NewInstallerFilename);
            if (File.Exists(newInstallerPath)) {
                File.Delete(newInstallerPath);
            }
        }

        private void AdjustRegistryForInstaller(InstallInformation information) {
            try {
                var keyName = $@"SOFTWARE\Classes\Installer\Products\{information.RegistryCode}\SourceList";
                var sourceListKey = Registry.LocalMachine.OpenSubKey(keyName, RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (sourceListKey != null) {
                    // HKEY_CLASSES_ROOT\Installer\Products\08CEA38BD6D37D44B8AAC9326BCD0266\SourceList
                    // HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Installer\Products\08CEA38BD6D37D44B8AAC9326BCD0266\SourceList
                    if (sourceListKey.GetValue("LastUsedSource") is string) {
                        sourceListKey.SetValue("LastUsedSource", $"n;1;{information.InstallFolder}");
                    }
                    if (sourceListKey.GetValue("PackageName") is string) {
                        sourceListKey.SetValue("PackageName", NewInstallerFilename);
                    }
                    sourceListKey.Close();
                }
            } catch (Exception ex) {
                throw new Exception($"{ex} (source list)");
            }

            try {
                var keyName = $@"SOFTWARE\Classes\Installer\Products\{information.RegistryCode}\SourceList\Net";
                var netKey = Registry.LocalMachine.OpenSubKey(keyName, RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (netKey != null) {
                    // HKEY_CLASSES_ROOT\Installer\Products\08CEA38BD6D37D44B8AAC9326BCD0266\SourceList\Net
                    // HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Installer\Products\08CEA38BD6D37D44B8AAC9326BCD0266\SourceList\Net
                    if (netKey.GetValue("1") is string) {
                        netKey.SetValue("1", information.InstallFolder);
                    }
                    netKey.Close();
                }
            } catch (Exception ex) {
                throw new Exception($"{ex} (source list net)");
            }

            try {
                var keyName = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\{information.RegistryCode}\InstallProperties";
                var installerPropertiesKey = Registry.LocalMachine.OpenSubKey(keyName, RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (installerPropertiesKey != null) {
                    // HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\08CEA38BD6D37D44B8AAC9326BCD0266\InstallProperties
                    if (installerPropertiesKey.GetValue("InstallSource") is string) {
                        installerPropertiesKey.SetValue("InstallSource", information.InstallFolder);
                    }
                    installerPropertiesKey.DeleteValue("NoModify");
                    installerPropertiesKey.Close();
                }
            } catch (Exception ex) {
                throw new Exception($"{ex} (install properties)");
            }

            try {
                var keyName = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{information.ProductCode}";
                var uninstallKey = Registry.LocalMachine.OpenSubKey(keyName, RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (uninstallKey != null) {
                    // HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{B83AEC80-3D6D-44D7-8BAA-9C23B6DC2066}
                    if (uninstallKey.GetValue("InstallSource") is string) {
                        uninstallKey.SetValue("InstallSource", information.InstallFolder);
                    }
                    uninstallKey.DeleteValue("NoModify");
                    uninstallKey.Close();
                }
            } catch (Exception ex) {
                throw new Exception($"{ex} (uninstall information)");
            }
        }

        #endregion

    }

}
