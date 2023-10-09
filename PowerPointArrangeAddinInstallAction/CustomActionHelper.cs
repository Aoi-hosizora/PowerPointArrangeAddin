using System;
using System.IO;
using System.Xml;
using Microsoft.Win32;

#nullable enable

namespace PowerPointArrangeAddinInstallAction {

    class CustomActionHelper {

        public CustomActionHelper(string? installFolder) {
            _installFolder = installFolder ?? throw new NullReferenceException(nameof(installFolder));
        }

        private readonly string _installFolder;

        public void RegisterAddIn() {
            AddInRegistryEntry entry;
            try {
                entry = GetAddInRegistryEntry();
            } catch (Exception ex) {
                throw new Exception($"Failed to read manifest file:\r\n\r\n{ex.Message}");
            }

            try {
                UpdateRegistry(entry);
            } catch (Exception ex) {
                throw new Exception($"Failed to update registry:\r\n\r\n{ex.Message}");
            }
        }

        public void UnregisterAddIn() {
            AddInRegistryEntry entry;
            try {
                entry = GetAddInRegistryEntry();
            } catch (Exception ex) {
                throw new Exception($"Failed to read manifest file:\r\n\r\n{ex.Message}");
            }

            try {
                DeleteRegistry(entry);
            } catch (Exception ex) {
                throw new Exception($"Failed to update registry:\r\n\r\n{ex.Message}");
            }
        }

        private string GetVstoFilePath() {
            return Path.Combine(_installFolder, "PowerPointArrangeAddin.vsto");
        }

        private string GetManifestFilePath() {
            return Path.Combine(_installFolder, "PowerPointArrangeAddin.dll.manifest");
        }

        private struct AddInRegistryEntry {
            public string DllName { get; set; }
            public string FriendlyName { get; set; }
            public string Description { get; set; }
            public int LoadBehavior { get; set; }
            public string ManifestPath { get; set; }
        }

        private AddInRegistryEntry GetAddInRegistryEntry() {
            var manifestPath = GetManifestFilePath();
            var doc = new XmlDocument();
            doc.Load(manifestPath); // allow throwing here

            var entry = new AddInRegistryEntry {
                DllName = "",
                FriendlyName = "",
                Description = "",
                LoadBehavior = 0,
                ManifestPath = $"file://{GetVstoFilePath().Replace("\\", "/")}|vstolocal"
            };

            var asmv1 = "urn:schemas-microsoft-com:asm.v1";
            var vstov4 = "urn:schemas-microsoft-com:vsto.v4";

            var assemblyIdentityElements = doc.GetElementsByTagName("assemblyIdentity", asmv1);
            if (assemblyIdentityElements.Count > 0) {
                entry.DllName = assemblyIdentityElements[0].Attributes?["name"]?.Value ?? "";
                if (entry.DllName.EndsWith(".dll")) {
                    entry.DllName = entry.DllName.Substring(0, entry.DllName.Length - 4);
                }
            }

            var friendlyNameElements = doc.GetElementsByTagName("friendlyName", vstov4);
            if (friendlyNameElements.Count > 0) {
                entry.FriendlyName = friendlyNameElements[0].InnerText;
            }

            var descriptionElements = doc.GetElementsByTagName("description", vstov4);
            if (descriptionElements.Count > 0) {
                entry.Description = descriptionElements[0].InnerText;
            }

            var appAddInElements = doc.GetElementsByTagName("appAddIn", vstov4);
            if (appAddInElements.Count > 0) {
                if (int.TryParse(appAddInElements[0].Attributes?["loadBehavior"]?.Value ?? "0", out var value)) {
                    entry.LoadBehavior = value;
                }
            }

            return entry;
        }

        private RegistryKey ThisRegistryKey => Registry.LocalMachine;

        private void UpdateRegistry(AddInRegistryEntry entry) {
            var addinsKeyPath = @"SOFTWARE\Microsoft\Office\PowerPoint\Addins";
            var addinsKey = ThisRegistryKey.OpenSubKey(addinsKeyPath, RegistryKeyPermissionCheck.ReadWriteSubTree);
            addinsKey ??= ThisRegistryKey.CreateSubKey(addinsKeyPath, RegistryKeyPermissionCheck.ReadWriteSubTree);
            if (addinsKey == null) {
                throw new NullReferenceException(nameof(addinsKey));
            }

            var addinKeyName = $"AoiHosizora.{entry.DllName}";
            var addinKey = addinsKey.OpenSubKey(addinKeyName);
            if (addinKey != null) {
                addinsKey.DeleteSubKeyTree(addinKeyName);
            }
            addinKey = addinsKey.CreateSubKey(addinKeyName);
            if (addinKey == null) {
                throw new NullReferenceException(nameof(addinKey));
            }

            addinKey.SetValue("Description", entry.Description, RegistryValueKind.String);
            addinKey.SetValue("FriendlyName", entry.FriendlyName, RegistryValueKind.String);
            addinKey.SetValue("LoadBehavior", entry.LoadBehavior, RegistryValueKind.DWord);
            addinKey.SetValue("Manifest", entry.ManifestPath, RegistryValueKind.String);

            addinKey.Close();
            addinsKey.Close();
        }

        private void DeleteRegistry(AddInRegistryEntry entry) {
            var addinsKeyPath = @"SOFTWARE\Microsoft\Office\PowerPoint\Addins";
            var addinsKey = ThisRegistryKey.OpenSubKey(addinsKeyPath, RegistryKeyPermissionCheck.ReadWriteSubTree);
            if (addinsKey == null) {
                return;
            }

            var addinKeyName = $"AoiHosizora.{entry.DllName}";
            try {
                addinsKey.DeleteSubKeyTree(addinKeyName);
            } catch (Exception) {
                // ignored
            } finally {
                addinsKey.Close();
            }
        }

    }

}
