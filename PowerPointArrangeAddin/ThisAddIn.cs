using System;
using Office = Microsoft.Office.Core;
using PowerPointArrangeAddin.Misc;

#nullable enable

namespace PowerPointArrangeAddin {

    public partial class ThisAddIn {

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            // enable WinForms visual styles
            System.Windows.Forms.Application.EnableVisualStyles();

            // load add-in setting
            AddInSetting.Instance.Load();

            // localize add-in
            var defaultLanguageId = Application.LanguageSettings.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI];
            AddInLanguageChanger.RegisterAddIn(defaultLanguageId: defaultLanguageId, uiInvalidator: () => _ribbon?.UpdateElementUiAndInvalidateRibbon());
            AddInLanguageChanger.ChangeLanguage(AddInSetting.Instance.Language);

            // callback for ribbon controls status
            Application.WindowSelectionChange += _ => _ribbon?.InvalidateRibbon();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        private Ribbon.ArrangeRibbon? _ribbon;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            _ribbon ??= new Ribbon.ArrangeRibbon();
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
