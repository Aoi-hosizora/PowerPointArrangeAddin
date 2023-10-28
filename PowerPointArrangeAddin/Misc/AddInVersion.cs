using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

#nullable enable

namespace PowerPointArrangeAddin.Misc {

    public class AddInVersion {

        private AddInVersion() { }

        private static AddInVersion? _instance;

        public static AddInVersion Instance {
            get {
                _instance ??= new AddInVersion();
                return _instance;
            }
        }

        #region Assembly Version and Parsing Related

        public (int Major, int Minor, int Build) GetAssemblyVersion() {
            var ver = Assembly.GetExecutingAssembly().GetName().Version;
            return (ver.Major, ver.Minor, ver.Build);
        }

        public string GetAssemblyVersionInString() {
            var (major, minor, build) = GetAssemblyVersion();
            return $"{major}.{minor}.{build}";
        }

        private (int Major, int Minor, int Build)? ParseVersionString(string version) {
            var parts = version.Trim().TrimStart('v', 'V') // [vV]123.456.789
                .Split('.')
                .Select(s => (Success: int.TryParse(s, out var number), Number: number))
                .Where(t => t.Success && t.Number >= 0)
                .Select(t => t.Number)
                .ToArray();
            if (parts.Length != 3) {
                return null;
            }
            return (parts[0], parts[1], parts[2]);
        }

        private bool CompareIfLessThan((int Major, int Minor, int Build) a, (int Major, int Minor, int Build) b) {
            if (a.Major < b.Major) {
                return true;
            }
            if (a.Major == b.Major && a.Minor < b.Minor) {
                return true;
            }
            if (a.Major == b.Major && a.Minor == b.Minor && a.Build < b.Build) {
                return true;
            }
            return false;
        }

        #endregion

        #region AppCenter Request Related

        private const string AppCenterUrl = "https://install.appcenter.ms/users/aoihosizora/apps/powerpointarrangeaddin/distribution_groups/public";
        private const string GitHubReleaseUrl = "https://github.com/Aoi-hosizora/PowerPointArrangeAddin/releases";

        private const string AppSecret = "38c9b3db-88af-4d40-a0ac-defcf9d5466e";
        private const string DistributionGroupId = "4f87083d-c9fc-4e49-b5df-92b5235015b5";

        [JsonObject]
        public class ReleaseInformation {
            [JsonRequired] public string Version { get; set; } = "";
            public string ReleaseNotes { get; set; } = "";
            public string Fingerprint { get; set; } = "";
            public string DownloadUrl { get; set; } = "";
        }

        private async Task<ReleaseInformation> GetLatestReleaseInformation() {
            ServicePointManager.SecurityProtocol |=
                SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

            // https://appcenter.ms/users/AoiHosizora/apps/PowerPointArrangeAddin/distribute/releases
            using var client = new HttpClient();
            var url = $"https://api.appcenter.ms/v0.1/public/sdk/apps/{AppSecret}/distribution_groups/{DistributionGroupId}/releases/latest";

            try {
                var response = await client.GetAsync(url);
                var body = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode) {
                    var errorMessage = JsonConvert.DeserializeObject<Dictionary<string, string>>(body)?["message"]?.Trim();
                    errorMessage ??= !string.IsNullOrWhiteSpace(errorMessage) ? errorMessage : response.StatusCode.ToString();
                    throw new Exception(errorMessage);
                }

                ReleaseInformation information;
                try {
                    var contractResolver = new DefaultContractResolver { NamingStrategy = new SnakeCaseNamingStrategy() };
                    var serializationSetting = new JsonSerializerSettings { ContractResolver = contractResolver };
                    information = JsonConvert.DeserializeObject<ReleaseInformation>(body, serializationSetting)!;
                } catch (Exception) {
                    throw new Exception("invalid response");
                }

                return information;
            } catch (Exception ex) {
                throw new Exception(ex.Message);
            }
        }

        #endregion

        #region Check Update Related

        public class CheckUpdateOptions {
            public bool ShowDialogForUpdates { get; init; } = true;
            public bool ShowDialogIfNoUpdate { get; init; }
            public bool ShowCheckingDialog { get; init; }
            public bool ShowDialogWhenException { get; init; }
            public bool ShowMoreOptionsForAutoCheck { get; init; }
            public IntPtr Owner { get; init; } = IntPtr.Zero;
        }

        public async Task<ReleaseInformation?> CheckUpdate(CheckUpdateOptions? options = null) {
            options ??= new CheckUpdateOptions();

            if (options.ShowMoreOptionsForAutoCheck && CheckIfIgnoreUpdate(null)) {
                return null; // ignore by date
            }

            Action? closeProgress = null;
            CancellationToken? cancellationToken = null;
            if (options.ShowCheckingDialog) {
                (closeProgress, cancellationToken) = ProgressDialog(options.Owner);
            }

            ReleaseInformation information;
            try {
                information = await GetLatestReleaseInformation();
            } catch (Exception ex) {
                if (options.ShowDialogWhenException) {
                    MessageBox.Show($"{MiscResources.Dlg_GetReleaseFailedText}\r\n{ex.Message}");
                }
                return null;
            } finally {
                closeProgress?.Invoke();
            }
            if (cancellationToken?.IsCancellationRequested == true) {
                return null;
            }

            var latestVersion = ParseVersionString(information.Version);
            if (latestVersion == null) {
                return null;
            }

            var currentVersion = GetAssemblyVersion();
            if (!CompareIfLessThan(currentVersion, latestVersion.Value)) {
                if (options.ShowDialogIfNoUpdate) {
                    NoUpdateDialog(options.Owner);
                }
                return null;
            }

            if (options.ShowMoreOptionsForAutoCheck && CheckIfIgnoreUpdate(information.Version)) {
                return null; // ignore by version
            }

            if (options.ShowDialogForUpdates) {
                HasUpdateDialog(information, options);
            }
            return information;
        }

        #endregion

        #region Check Update Dialog Related

        private (Action CloseDialog, CancellationToken CancellationToken) ProgressDialog(IntPtr owner) {
            TaskDialog dialog;
            var cts = new CancellationTokenSource();
            using (new EnableThemingInScope(true)) {
                dialog = new TaskDialog {
                    Caption = AddInDescription.Instance.Title,
                    Icon = TaskDialogStandardIcon.Information,
                    InstructionText = MiscResources.Dlg_CheckUpdateProgressText,
                    ProgressBar = new TaskDialogProgressBar { State = TaskDialogProgressBarState.Marquee },
                    OwnerWindowHandle = owner,
                    StandardButtons = TaskDialogStandardButtons.Cancel
                };
                new Thread(() => {
                    var result = dialog.Show();
                    if (result == TaskDialogResult.Cancel) {
                        cts.Cancel();
                    } else if (result == TaskDialogResult.Close) {
                        // ignored
                    }
                }).Start();
            }

            void Close() {
                try {
                    dialog.Close();
                } catch (Exception) {
                    // may already closed, ignored
                }
            }

            return (CloseDialog: Close, CancellationToken: cts.Token);
        }

        private void NoUpdateDialog(IntPtr owner) {
            using (new EnableThemingInScope(true)) {
                var dialog = new TaskDialog {
                    Caption = AddInDescription.Instance.Title,
                    Icon = TaskDialogStandardIcon.Information,
                    InstructionText = MiscResources.Dlg_NoUpdateText,
                    Text = string.Format(MiscResources.Dlg_CurrentIsNewestVersionText, $"v{GetAssemblyVersionInString()}"),
                    OwnerWindowHandle = owner,
                    StandardButtons = TaskDialogStandardButtons.Ok
                };
                dialog.Show();
            }
        }

        private void HasUpdateDialog(ReleaseInformation information, CheckUpdateOptions options) {
            using (new EnableThemingInScope(true)) {
                var dialog = new TaskDialog {
                    Caption = AddInDescription.Instance.Title,
                    Icon = TaskDialogStandardIcon.Information,
                    InstructionText = string.Format(MiscResources.Dlg_HasNewVersionReleasedText, $"v{information.Version}"),
                    Text = $"{string.Format(MiscResources.Dlg_CurrentVersionText, $"v{GetAssemblyVersionInString()}")}\r\n\r\n{MiscResources.Dlg_DownloadNewVersionQuestionText}",
                    DetailsExpandedText = $"{MiscResources.Dlg_ReleaseNoteText}\r\n\r\n{information.ReleaseNotes}",
                    ExpansionMode = TaskDialogExpandedDetailsLocation.ExpandFooter,
                    DetailsExpanded = false,
                    OwnerWindowHandle = options.Owner,
                    StandardButtons = TaskDialogStandardButtons.Cancel
                };

                var lnkAppCenter = new TaskDialogCommandLink("AppCenter", MiscResources.Dlg_VisitAppCenterText);
                var lnkGitHub = new TaskDialogCommandLink("GitHub", MiscResources.Dlg_VisitGitHubText);
                lnkAppCenter.Click += (_, _) => Process.Start(AppCenterUrl);
                lnkGitHub.Click += (_, _) => Process.Start(GitHubReleaseUrl);
                dialog.Controls.Add(lnkAppCenter);
                dialog.Controls.Add(lnkGitHub);

                if (options.ShowMoreOptionsForAutoCheck) {
                    var lnkIgnoreUntil = new TaskDialogCommandLink("Ignore until", MiscResources.Dlg_IgnoreUntilTomorrowText);
                    var lnkIgnoreVersion = new TaskDialogCommandLink("Ignore version", MiscResources.Dlg_IgnoreThisVersion);
                    var lnkDisableAutoCheck = new TaskDialogCommandLink("Disable auto check", MiscResources.Dlg_DisableAutoCheckUpdateText);
                    lnkIgnoreUntil.Click += (_, _) => IgnoreSpecificUpdate(null, () => dialog.Close());
                    lnkIgnoreVersion.Click += (_, _) => IgnoreSpecificUpdate(information.Version, () => dialog.Close());
                    lnkDisableAutoCheck.Click += (_, _) => {
                        AddInSetting.Instance.CheckUpdateWhenStartUp = false;
                        AddInSetting.Instance.Save();
                        dialog.Close();
                    };
                    dialog.Controls.Add(lnkIgnoreUntil);
                    dialog.Controls.Add(lnkIgnoreVersion);
                    dialog.Controls.Add(lnkDisableAutoCheck);
                }

                dialog.Show();
            }
        }

        private void IgnoreSpecificUpdate(string? version, Action? postAction = null) {
            AddInSetting.Instance.IgnoreUpdateRecord = version == null
                ? $"date={DateTime.Today:yyyyMMdd}"
                : $"version={version}";
            AddInSetting.Instance.Save();
            postAction?.Invoke();
        }

        private bool CheckIfIgnoreUpdate(string? version) {
            var record = AddInSetting.Instance.IgnoreUpdateRecord.Trim();
            if (record == $"date={DateTime.Today:yyyyMMdd}") {
                return true;
            }
            if (version != null && record == $"version={version}") {
                return true;
            }
            return false;
        }

        #endregion

    }

}