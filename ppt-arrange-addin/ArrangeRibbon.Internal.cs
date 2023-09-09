using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace ppt_arrange_addin {

    using RES = Properties.Resources;
    using ARES = ArrangeRibbonResources;

    [ComVisible(true)]
    public partial class ArrangeRibbon : Office.IRibbonExtensibility {

        public ArrangeRibbon() {
            InitializeAvailabilityRules();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId) {
            return GetResourceText("ppt_arrange_addin.ArrangeRibbon.xml");
        }

        private static string GetResourceText(string resourceName) {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            foreach (var name in resourceNames) {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0) {
                    var stream = asm.GetManifestResourceStream(name);
                    if (stream != null) {
                        using var resourceReader = new StreamReader(stream);
                        return resourceReader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        #endregion

        #region UI Callbacks For Ribbon Xml

        private class ElementUi {
            public string Label { get; init; }
            public System.Drawing.Image Image { get; init; }
        }

        // ReSharper disable InconsistentNaming
        private const string grpArrange = "grpArrange";
        private const string btnAlignLeft = "btnAlignLeft";
        private const string btnAlignCenter = "btnAlignCenter";
        private const string btnAlignRight = "btnAlignRight";
        private const string btnAlignTop = "btnAlignTop";
        private const string btnAlignMiddle = "btnAlignMiddle";
        private const string btnAlignBottom = "btnAlignBottom";
        private const string btnDistributeHorizontal = "btnDistributeHorizontal";
        private const string btnDistributeVertical = "btnDistributeVertical";
        private const string btnScaleSameWidth = "btnScaleSameWidth";
        private const string btnScaleSameHeight = "btnScaleSameHeight";
        private const string btnScaleSameSize = "btnScaleSameSize";
        private const string btnScalePosition = "btnScalePosition";
        private const string btnExtendSameLeft = "btnExtendSameLeft";
        private const string btnExtendSameRight = "btnExtendSameRight";
        private const string btnExtendSameTop = "btnExtendSameTop";
        private const string btnExtendSameBottom = "btnExtendSameBottom";
        private const string btnSnapLeft = "btnSnapLeft";
        private const string btnSnapRight = "btnSnapRight";
        private const string btnSnapTop = "btnSnapTop";
        private const string btnSnapBottom = "btnSnapBottom";
        private const string btnMoveForward = "btnMoveForward";
        private const string btnMoveFront = "btnMoveFront";
        private const string btnMoveBackward = "btnMoveBackward";
        private const string btnMoveBack = "btnMoveBack";
        private const string btnRotateRight90 = "btnRotateRight90";
        private const string btnRotateLeft90 = "btnRotateLeft90";
        private const string btnFlipVertical = "btnFlipVertical";
        private const string btnFlipHorizontal = "btnFlipHorizontal";
        private const string btnGroup = "btnGroup";
        private const string btnUngroup = "btnUngroup";
        private const string grpShapePosition = "grpShapePosition";
        private const string edtShapePositionX = "edtShapePositionX";
        private const string edtShapePositionY = "edtShapePositionY";
        private const string btnShapePositionCopy = "btnShapePositionCopy";
        private const string btnShapePositionPaste = "btnShapePositionPaste";
        private const string grpTextbox = "grpTextbox";
        private const string btnAutofitOff = "btnAutofitOff";
        private const string btnAutofitText = "btnAutofitText";
        private const string btnAutoResize = "btnAutoResize";
        private const string btnWrapText = "btnWrapText";
        private const string edtMarginLeft = "edtMarginLeft";
        private const string edtMarginRight = "edtMarginRight";
        private const string edtMarginTop = "edtMarginTop";
        private const string edtMarginBottom = "edtMarginBottom";
        private const string btnResetMarginHorizontal = "btnResetMarginHorizontal";
        private const string btnResetMarginVertical = "btnResetMarginVertical";
        // ReSharper restore InconsistentNaming

        private readonly Dictionary<string, ElementUi> _elementLabels = new() {
            { grpArrange, new ElementUi { Label = ARES.grpArrange, Image = RES.ObjectSendToBack } },
            { btnAlignLeft, new ElementUi { Label = ARES.btnAlignLeft, Image = RES.ObjectsAlignLeft } },
            { btnAlignCenter, new ElementUi { Label = ARES.btnAlignCenter, Image = RES.ObjectsAlignCenterHorizontal } },
            { btnAlignRight, new ElementUi { Label = ARES.btnAlignRight, Image = RES.ObjectsAlignRight } },
            { btnAlignTop, new ElementUi { Label = ARES.btnAlignTop, Image = RES.ObjectsAlignTop } },
            { btnAlignMiddle, new ElementUi { Label = ARES.btnAlignMiddle, Image = RES.ObjectsAlignMiddleVertical } },
            { btnAlignBottom, new ElementUi { Label = ARES.btnAlignBottom, Image = RES.ObjectsAlignBottom } },
            { btnDistributeHorizontal, new ElementUi { Label = ARES.btnDistributeHorizontal, Image = RES.AlignDistributeHorizontally } },
            { btnDistributeVertical, new ElementUi { Label = ARES.btnDistributeVertical, Image = RES.AlignDistributeVertically } },
            { btnScaleSameWidth, new ElementUi { Label = ARES.btnScaleSameWidth, Image = RES.ScaleSameWidth } },
            { btnScaleSameHeight, new ElementUi { Label = ARES.btnScaleSameHeight, Image = RES.ScaleSameHeight } },
            { btnScaleSameSize, new ElementUi { Label = ARES.btnScaleSameSize, Image = RES.ScaleSameSize } },
            { btnScalePosition, new ElementUi { Label = ARES.btnScalePosition_Middle, Image = RES.ScaleFromMiddle } },
            { btnExtendSameLeft, new ElementUi { Label = ARES.btnExtendSameLeft, Image = RES.ExtendSameLeft } },
            { btnExtendSameRight, new ElementUi { Label = ARES.btnExtendSameRight, Image = RES.ExtendSameRight } },
            { btnExtendSameTop, new ElementUi { Label = ARES.btnExtendSameTop, Image = RES.ExtendSameTop } },
            { btnExtendSameBottom, new ElementUi { Label = ARES.btnExtendSameBottom, Image = RES.ExtendSameBottom } },
            { btnSnapLeft, new ElementUi { Label = ARES.btnSnapLeft, Image = RES.SnapToLeft } },
            { btnSnapRight, new ElementUi { Label = ARES.btnSnapRight, Image = RES.SnapToRight } },
            { btnSnapTop, new ElementUi { Label = ARES.btnSnapTop, Image = RES.SnapToTop } },
            { btnSnapBottom, new ElementUi { Label = ARES.btnSnapBottom, Image = RES.SnapToBottom } },
            { btnMoveForward, new ElementUi { Label = ARES.btnMoveForward, Image = RES.ObjectBringForward } },
            { btnMoveFront, new ElementUi { Label = ARES.btnMoveFront, Image = RES.ObjectBringToFront } },
            { btnMoveBackward, new ElementUi { Label = ARES.btnMoveBackward, Image = RES.ObjectSendBackward } },
            { btnMoveBack, new ElementUi { Label = ARES.btnMoveBack, Image = RES.ObjectSendToBack } },
            { btnRotateRight90, new ElementUi { Label = ARES.btnRotateRight90, Image = RES.ObjectRotateRight90 } },
            { btnRotateLeft90, new ElementUi { Label = ARES.btnRotateLeft90, Image = RES.ObjectRotateLeft90 } },
            { btnFlipVertical, new ElementUi { Label = ARES.btnFlipVertical, Image = RES.ObjectFlipVertical } },
            { btnFlipHorizontal, new ElementUi { Label = ARES.btnFlipHorizontal, Image = RES.ObjectFlipHorizontal } },
            { btnGroup, new ElementUi { Label = ARES.btnGroup, Image = RES.ObjectsGroup } },
            { btnUngroup, new ElementUi { Label = ARES.btnUngroup, Image = RES.ObjectsUngroup } },
            { grpShapePosition, new ElementUi { Label = ARES.grpShapePosition, Image = RES.ObjectPosition } },
            { edtShapePositionX, new ElementUi { Label = ARES.edtShapePositionX } },
            { edtShapePositionY, new ElementUi { Label = ARES.edtShapePositionY } },
            { btnShapePositionCopy, new ElementUi { Label = ARES.btnShapePositionCopy, Image = RES.Copy } },
            { btnShapePositionPaste, new ElementUi { Label = ARES.btnShapePositionPaste, Image = RES.Paste } },
            { grpTextbox, new ElementUi { Label = ARES.grpTextbox, Image = RES.TextboxWrapText } },
            { btnAutofitOff, new ElementUi { Label = ARES.btnAutofitOff, Image = RES.TextboxAutofitOff } },
            { btnAutofitText, new ElementUi { Label = ARES.btnAutofitText, Image = RES.TextboxAutofitText } },
            { btnAutoResize, new ElementUi { Label = ARES.btnAutoResize, Image = RES.TextboxAutoResize } },
            { btnWrapText, new ElementUi { Label = ARES.btnWrapText, Image = RES.TextboxWrapText_32 } },
            { edtMarginLeft, new ElementUi { Label = ARES.edtMarginLeft } },
            { edtMarginRight, new ElementUi { Label = ARES.edtMarginRight } },
            { edtMarginTop, new ElementUi { Label = ARES.edtMarginTop } },
            { edtMarginBottom, new ElementUi { Label = ARES.edtMarginBottom } },
            { btnResetMarginHorizontal, new ElementUi { Label = ARES.btnResetMarginHorizontal, Image = RES.TextboxResetMargin } },
            { btnResetMarginVertical, new ElementUi { Label = ARES.btnResetMarginVertical, Image = RES.TextboxResetMargin } },
        };

        public string GetLabel(Office.IRibbonControl ribbonControl) {
            _elementLabels.TryGetValue(ribbonControl.Id, out var eui);
            return eui?.Label ?? "<Unknown>";
        }

        public System.Drawing.Image GetImage(Office.IRibbonControl ribbonControl) {
            _elementLabels.TryGetValue(ribbonControl.Id, out var eui);
            return eui?.Image;
        }

        #endregion

    }

}
