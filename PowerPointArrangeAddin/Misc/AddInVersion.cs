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

        public (int Major, int Minor, int Build) GetAssemblyVersion() {
            var ver = Assembly.GetExecutingAssembly().GetName().Version;
            return (ver.Major, ver.Minor, ver.Build);
        }

        public string GetAssemblyVersionInString() {
            var (major, minor, build) = GetAssemblyVersion();
            return $"{major}.{minor}.{build}";
        }

        private (int Major, int Minor, int Build)? ParseVersionString(string version) {
            var parts = version.TrimStart('v', 'V') // [vV]000.000.000
                .Split('.')
                .Select(s => (Success: int.TryParse(s, out var number), Number: number))
                .Where(t => t.Success && t.Number >= 0)
                .Select(t => t.Number)
                .ToArray();
            if (parts.Length != 3) {
                return null; // invalid version string
            }
            return (parts[0], parts[1], parts[2]);
        }

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

        private async Task<ReleaseInformation> QueryLatestReleaseInformation() {
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
                throw new Exception($"Failed to get release information: \r\n{ex.Message}");
            }
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

        public class CheckUpdateOptions {
            public bool ShowDialogForUpdates { get; init; } = true;
            public bool ShowDialogIfNoUpdate { get; init; }
            public bool ShowCheckingDialog { get; init; }
            public bool ShowDialogWhenException { get; init; }
            public IntPtr Owner { get; init; } = IntPtr.Zero;
        }

        public async Task<ReleaseInformation?> CheckUpdate(CheckUpdateOptions? options = null) {
            options ??= new CheckUpdateOptions();

            TaskDialog? pgDialog = null;
            var cts = new CancellationTokenSource();
            if (options.ShowCheckingDialog) {
                using (new EnableThemingInScope(true)) {
                    pgDialog = new TaskDialog {
                        Caption = AddInDescription.Instance.Title,
                        InstructionText = "Checking for updates...",
                        Icon = TaskDialogStandardIcon.Information,
                        ProgressBar = new TaskDialogProgressBar { State = TaskDialogProgressBarState.Marquee },
                        Cancelable = false,
                        OwnerWindowHandle = options.Owner,
                        StandardButtons = TaskDialogStandardButtons.Cancel,
                    };
                    new Thread(() => {
                        var result = pgDialog.Show();
                        if (result == TaskDialogResult.Cancel) {
                            cts.Cancel();
                        }
                    }).Start();
                }
            }

            ReleaseInformation information;
            try {
                information = await QueryLatestReleaseInformation();
            } catch (Exception ex) {
                if (options.ShowDialogWhenException) {
                    MessageBox.Show(ex.Message);
                }
                return null;
            } finally {
                try {
                    pgDialog?.Close();
                } catch (Exception) { }
            }
            if (cts.IsCancellationRequested) {
                return null;
            }

            var latestVersion = ParseVersionString(information.Version);
            if (latestVersion == null) {
                return null; // invalid version string
            }

            var currentVersion = GetAssemblyVersion();
            if (!CompareIfLessThan(currentVersion, latestVersion.Value)) {
                if (options.ShowDialogIfNoUpdate) {
                    using (new EnableThemingInScope(true)) {
                        var dialog = new TaskDialog {
                            Caption = AddInDescription.Instance.Title,
                            InstructionText = "There are currently no updates available.",
                            Text = $"Current version (v{GetAssemblyVersionInString()}) is the newest version!",
                            Icon = TaskDialogStandardIcon.Information,
                            Cancelable = false,
                            OwnerWindowHandle = options.Owner,
                            StandardButtons = TaskDialogStandardButtons.Ok
                        };
                        dialog.Show();
                    }
                }
                return null;
            }

            if (options.ShowDialogForUpdates) {
                using (new EnableThemingInScope(true)) {
                    var dialog = new TaskDialog();

                    dialog.Caption = AddInDescription.Instance.Title;
                    dialog.InstructionText = $"v{information.Version} has been released!";
                    dialog.Text = $"Current version is v{GetAssemblyVersionInString()}.\r\nDo you want to download the new version?";
                    dialog.Icon = TaskDialogStandardIcon.Information;

                    dialog.DetailsExpandedText = $"\r\n{information.ReleaseNotes}";
                    dialog.ExpansionMode = TaskDialogExpandedDetailsLocation.ExpandContent;
                    dialog.DetailsExpanded = false;

                    dialog.Cancelable = false;
                    dialog.OwnerWindowHandle = options.Owner;
                    dialog.StandardButtons = TaskDialogStandardButtons.Cancel;

                    var lnkAppCenter = new TaskDialogCommandLink("AppCenter", "Visit &AppCenter to download");
                    var lnkGitHub = new TaskDialogCommandLink("GitHub", "Visit &GitHub to download");
                    lnkAppCenter.Click += (_, _) => Process.Start(AppCenterUrl);
                    lnkGitHub.Click += (_, _) => Process.Start(GitHubReleaseUrl);
                    dialog.Controls.Add(lnkAppCenter);
                    dialog.Controls.Add(lnkGitHub);

                    dialog.Show();
                }
            }
            return information;
        }

    }

}
