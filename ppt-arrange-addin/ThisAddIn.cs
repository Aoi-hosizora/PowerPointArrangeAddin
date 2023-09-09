using System;
using Office = Microsoft.Office.Core;
using System.Threading;
using System.Globalization;

namespace ppt_arrange_addin {

    public partial class ThisAddIn {

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            // localized add-in
            var languageId = Application.LanguageSettings.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI];
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(languageId);
            Properties.Resources.Culture = new CultureInfo(languageId);
            ArrangeRibbonResources.Culture = new CultureInfo(languageId); // TODO zh-CN

            // ribbon controls status
            Application.WindowSelectionChange += _ => {
                _ribbon.AdjustRibbonButtonsAvailability();
            };
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        private ArrangeRibbon _ribbon;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            _ribbon ??= new ArrangeRibbon();
            return _ribbon;
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
