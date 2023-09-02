using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace ppt_arrange_addin {

    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility {

        public Ribbon() { }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("ppt_arrange_addin.Ribbon.xml");
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
            ribbon = ribbonUI;
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
