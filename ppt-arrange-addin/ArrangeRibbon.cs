using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ppt_arrange_addin {

    [ComVisible(true)]
    public class ArrangeRibbon : Office.IRibbonExtensibility {

        public ArrangeRibbon() { }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("ppt_arrange_addin.ArrangeRibbon.xml");
        }

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i) {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if (resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        private Office.IRibbonUI ribbon;

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            this.ribbon = ribbonUI;
        }
        private PowerPoint.ShapeRange GetShapeRange(int mustMoreThanOrEqualTo = 1) {
            var shapeRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            if (shapeRange.Count < mustMoreThanOrEqualTo) {
                return null;
            }
            return shapeRange;
        }

        private void StartNewUndoEntry() {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
        }

        public void BtnAlignLeft_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            StartNewUndoEntry();
            var flag = shapeRange.Count == 1 ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            shapeRange.Align(Office.MsoAlignCmd.msoAlignLefts, flag);
        }

        public System.Drawing.Image GetImage(string ImageName) {
            return Properties.Resources.ResourceManager.GetObject(ImageName) as System.Drawing.Image;
        }
    }
}
