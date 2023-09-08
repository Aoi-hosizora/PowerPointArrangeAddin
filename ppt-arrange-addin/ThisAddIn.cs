using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Threading;
using System.Globalization;

namespace ppt_arrange_addin {

    public partial class ThisAddIn {

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            // localized add-in
            var languageId = Application.LanguageSettings.get_LanguageID(Office.MsoAppLanguageID.msoLanguageIDUI);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(languageId);
            Properties.Resources.Culture = new CultureInfo(languageId);
            ArrangeRibbonResources.Culture = new CultureInfo(languageId); // TODO zh-CN

            // ribbon controls status
            Application.WindowSelectionChange += (selection) => {
                ribbon.AdjustRibbonButtonsAvailability();
            };
            Application.AfterDragDropOnSlide += (slide, x, y) => {
                ribbon.AdjustRibbonButtonsAvailability(onlyForDrag: true);
            };
            Application.AfterShapeSizeChange += (shape) => {
                ribbon.AdjustRibbonButtonsAvailability(onlyForDrag: true);
            };
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        private ArrangeRibbon ribbon;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            if (ribbon == null) {
                ribbon = new ArrangeRibbon();
            }
            return ribbon;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

    }

}
