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
            PowerPoint.Application app = GetHostItem<PowerPoint.Application>(typeof(PowerPoint.Application), "Application");
            int lcid = app.LanguageSettings.get_LanguageID(Office.MsoAppLanguageID.msoLanguageIDUI);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(lcid);

            // ribbon accessibility
            Application.WindowSelectionChange += (obj) => {
                Globals.Ribbons.ArrangeRibbon.AdjustButtonsAccessibility();
            };
            Application.SlideSelectionChanged += (obj) => {
                Globals.Ribbons.ArrangeRibbon.AdjustButtonsAccessibility();
            };
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

    }

}
