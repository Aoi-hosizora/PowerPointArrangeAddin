using System;
using System.Windows.Forms;
using Microsoft.Deployment.WindowsInstaller;

#nullable enable

namespace PowerPointArrangeAddinInstallAction {

    public class CustomAction {

        [CustomAction]
        public static ActionResult RegisterAddIn(Session session) {
            session.Message(InstallMessage.Info, new Record { FormatString = "Register add-in" });
            try {
                var installFolder = session.CustomActionData["InstallFolder"];
                var helper = new CustomActionHelper(installFolder);
                helper.RegisterAddIn();
                return ActionResult.Success;
            } catch (Exception ex) {
                MsgBox(session, ex.Message, MessageBoxIcon.Error);
                return ActionResult.Failure;
            }
        }

        [CustomAction]
        public static ActionResult UnregisterAddIn(Session session) {
            session.Message(InstallMessage.Info, new Record { FormatString = "Unregister add-in" });
            try {
                var installFolder = session.CustomActionData["InstallFolder"];
                var helper = new CustomActionHelper(installFolder);
                helper.UnregisterAddIn();
                return ActionResult.Success;
            } catch (Exception ex) {
                MsgBox(session, ex.Message, MessageBoxIcon.Warning);
                return ActionResult.Success; // just return success rather than failure
            }
        }

        private static void MsgBox(Session session, string text, MessageBoxIcon icon) {
            var flag = InstallMessage.User + (int) icon + (int) MessageBoxButtons.OK;
            var record = new Record { FormatString = text };
            session.Message(flag, record);
        }

    }

}
