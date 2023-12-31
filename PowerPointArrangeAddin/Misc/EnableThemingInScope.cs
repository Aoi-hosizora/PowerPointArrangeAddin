using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Windows.Forms;

#nullable enable

// ReSharper disable InconsistentNaming
// ReSharper disable UseObjectOrCollectionInitializer

// C#: comctl32.dll version 6 in debugger
// https://stackoverflow.com/questions/1415270/c-comctl32-dll-version-6-in-debugger

namespace PowerPointArrangeAddin.Misc {

    [SuppressUnmanagedCodeSecurity]
    internal class EnableThemingInScope : IDisposable {

        // Private data
        private UIntPtr cookie;

        private static ACTCTX enableThemingActivationContext;

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")]
        private static IntPtr hActCtx;

        private static bool contextCreationSucceeded;

        public EnableThemingInScope(bool enable) {
            cookie = UIntPtr.Zero;
            if (enable && OSFeature.Feature.IsPresent(OSFeature.Themes)) {
                if (EnsureActivateContextCreated()) {
                    if (!ActivateActCtx(hActCtx, out cookie)) {
                        // Be sure cookie always zero if activation failed
                        cookie = UIntPtr.Zero;
                    }
                }
            }
        }

        ~EnableThemingInScope() {
            Dispose();
        }

        void IDisposable.Dispose() {
            Dispose();
            GC.SuppressFinalize(this);
        }

        private void Dispose() {
            if (cookie != UIntPtr.Zero) {
                try {
                    if (DeactivateActCtx(0, cookie)) {
                        // deactivation succeeded...
                        cookie = UIntPtr.Zero;
                    }
                } catch (SEHException) {
                    // Hopefully solved this exception
                }
            }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2002:DoNotLockOnObjectsWithWeakIdentity")]
        private static bool EnsureActivateContextCreated() {
            lock (typeof(EnableThemingInScope)) {
                if (!contextCreationSucceeded) {
                    // Pull manifest from the .NET Framework install
                    // directory

                    string? assemblyLoc;

                    var fiop = new FileIOPermission(PermissionState.None);
                    fiop.AllFiles = FileIOPermissionAccess.PathDiscovery;
                    fiop.Assert();
                    try {
                        assemblyLoc = typeof(object).Assembly.Location;
                    } finally {
                        CodeAccessPermission.RevertAssert();
                    }

                    string? manifestLoc = null;
                    string? installDir = null;
                    if (assemblyLoc != null) {
                        installDir = Path.GetDirectoryName(assemblyLoc);
                        const string manifestName = "XPThemes.manifest";
                        manifestLoc = Path.Combine(installDir!, manifestName);
                    }

                    if (manifestLoc != null && installDir != null) {
                        enableThemingActivationContext = new ACTCTX();
                        enableThemingActivationContext.cbSize = Marshal.SizeOf(typeof(ACTCTX));
                        enableThemingActivationContext.lpSource = manifestLoc;

                        // Set the lpAssemblyDirectory to the install
                        // directory to prevent Win32 Side by Side from
                        // looking for comctl32 in the application
                        // directory, which could cause a bogus dll to be
                        // placed there and open a security hole.
                        enableThemingActivationContext.lpAssemblyDirectory = installDir;
                        enableThemingActivationContext.dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID;

                        // Note this will fail gracefully if file specified
                        // by manifestLoc doesn't exist.
                        hActCtx = CreateActCtx(ref enableThemingActivationContext);
                        contextCreationSucceeded = (hActCtx != new IntPtr(-1));
                    }
                }

                // If we return false, we'll try again on the next call into
                // EnsureActivateContextCreated(), which is fine.
                return contextCreationSucceeded;
            }
        }

        // All the pinvoke goo...
        [DllImport("Kernel32.dll")]
        private static extern IntPtr CreateActCtx(ref ACTCTX actctx);

        [DllImport("Kernel32.dll")]
        private static extern bool ActivateActCtx(IntPtr hActCtx, out UIntPtr lpCookie);

        [DllImport("Kernel32.dll")]
        private static extern bool DeactivateActCtx(uint dwFlags, UIntPtr lpCookie);

        private const int ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID = 0x004;

#pragma warning disable CS0649
        // ReSharper disable NotAccessedField.Local
        private struct ACTCTX {
            public int cbSize;
            public uint dwFlags;
            public string lpSource;
            public ushort wProcessorArchitecture;
            public ushort wLangId;
            public string lpAssemblyDirectory;
            public string lpResourceName;
            public string lpApplicationName;
        }
        // ReSharper restore NotAccessedField.Local

    }

}
