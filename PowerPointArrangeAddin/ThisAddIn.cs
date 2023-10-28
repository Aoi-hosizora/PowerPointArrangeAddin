using System;
using Office = Microsoft.Office.Core;

#nullable enable

namespace PowerPointArrangeAddin {

    public partial class ThisAddIn {

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            // enable WinForms visual styles
            System.Windows.Forms.Application.EnableVisualStyles();

            // load add-in setting
            Misc.AddInSetting.Instance.Load();

            // localize add-in
            var defaultLanguageId = Application.LanguageSettings.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI];
            Misc.AddInLanguageChanger.RegisterAddIn(defaultLanguageId: defaultLanguageId);
            Misc.AddInLanguageChanger.ChangeLanguage(Misc.AddInSetting.Instance.Language);

            // callback for ribbon controls status
            Application.WindowSelectionChange += _ => Ribbon.ArrangeRibbon.Instance.InvalidateRibbon();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            return Ribbon.ArrangeRibbon.Instance;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion

    }

}
