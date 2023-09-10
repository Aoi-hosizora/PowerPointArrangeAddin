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

        #region Ribbon Elements ID

        // ReSharper disable InconsistentNaming
        // grpArrange
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
        // grpTextbox
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
        // grpShapeSizeAndPosition
        private const string grpShapeSizeAndPosition = "grpShapeSizeAndPosition";
        private const string mnuShapeArrangement = "mnuShapeArrangement";
        private const string btnShapeScalePosition = "btnShapeScalePosition";
        private const string btnLockShapeAspectRatio = "btnLockShapeAspectRatio";
        private const string btnCopyShapeSize = "btnCopyShapeSize";
        private const string btnPasteShapeSize = "btnPasteShapeSize";
        private const string edtShapePositionX = "edtShapePositionX";
        private const string edtShapePositionY = "edtShapePositionY";
        private const string btnCopyShapePosition = "btnCopyShapePosition";
        private const string btnPasteShapePosition = "btnPasteShapePosition";
        // grpReplacePicture
        private const string grpReplacePicture = "grpReplacePicture";
        private const string btnReplaceWithClipboard = "btnReplaceWithClipboard";
        private const string btnReplaceWithFile = "btnReplaceWithFile";
        private const string cbxReserveOriginalSize = "cbxReserveOriginalSize";
        private const string cbxReplaceToMiddle = "cbxReplaceToMiddle";
        // grpPictureSizeAndPosition
        private const string grpPictureSizeAndPosition = "grpPictureSizeAndPosition";
        private const string mnuPictureArrangement = "mnuPictureArrangement";
        private const string btnResetPictureSize = "btnResetPictureSize";
        private const string btnPictureScalePosition = "btnPictureScalePosition";
        private const string btnLockPictureAspectRatio = "btnLockPictureAspectRatio";
        private const string btnCopyPictureSize = "btnCopyPictureSize";
        private const string btnPastePictureSize = "btnPastePictureSize";
        private const string edtPicturePositionX = "edtPicturePositionX";
        private const string edtPicturePositionY = "edtPicturePositionY";
        private const string btnCopyPicturePosition = "btnCopyPicturePosition";
        private const string btnPastePicturePosition = "btnPastePicturePosition";
        // mnuArrangement
        private const string mnuArrangement_sepAlignmentAndResizing = "mnuArrangement_sepAlignmentAndResizing";
        private const string mnuArrangement_mnuAlignment = "mnuArrangement_mnuAlignment";
        private const string mnuArrangement_mnuResizing = "mnuArrangement_mnuResizing";
        private const string mnuArrangement_mnuSnapping = "mnuArrangement_mnuSnapping";
        private const string mnuArrangement_mnuRotation = "mnuArrangement_mnuRotation";
        private const string mnuArrangement_sepLayerOrderAndGrouping = "mnuArrangement_sepLayerOrderAndGrouping";
        private const string mnuArrangement_mnuLayerOrder = "mnuArrangement_mnuLayerOrder";
        private const string mnuArrangement_mnuGrouping = "mnuArrangement_mnuGrouping";
        private const string mnuArrangement_sepObjectsInSlide = "mnuArrangement_sepObjectsInSlide";
        // ReSharper restore InconsistentNaming

        #endregion

        #region Element Ui Callbacks

        private class ElementUi {
            public string Label { get; init; }
            public System.Drawing.Image Image { get; init; }
        }

        private readonly Dictionary<string, Func<ElementUi>> _elementLabels = new() {
            // grpArrange
            { grpArrange, () => new ElementUi { Label = ARES.grpArrange, Image = RES.ObjectArrangement } },
            { btnAlignLeft, () => new ElementUi { Label = ARES.btnAlignLeft, Image = RES.ObjectsAlignLeft } },
            { btnAlignCenter, () => new ElementUi { Label = ARES.btnAlignCenter, Image = RES.ObjectsAlignCenterHorizontal } },
            { btnAlignRight, () => new ElementUi { Label = ARES.btnAlignRight, Image = RES.ObjectsAlignRight } },
            { btnAlignTop, () => new ElementUi { Label = ARES.btnAlignTop, Image = RES.ObjectsAlignTop } },
            { btnAlignMiddle, () => new ElementUi { Label = ARES.btnAlignMiddle, Image = RES.ObjectsAlignMiddleVertical } },
            { btnAlignBottom, () => new ElementUi { Label = ARES.btnAlignBottom, Image = RES.ObjectsAlignBottom } },
            { btnDistributeHorizontal, () => new ElementUi { Label = ARES.btnDistributeHorizontal, Image = RES.AlignDistributeHorizontally } },
            { btnDistributeVertical, () => new ElementUi { Label = ARES.btnDistributeVertical, Image = RES.AlignDistributeVertically } },
            { btnScaleSameWidth, () => new ElementUi { Label = ARES.btnScaleSameWidth, Image = RES.ScaleSameWidth } },
            { btnScaleSameHeight, () => new ElementUi { Label = ARES.btnScaleSameHeight, Image = RES.ScaleSameHeight } },
            { btnScaleSameSize, () => new ElementUi { Label = ARES.btnScaleSameSize, Image = RES.ScaleSameSize } },
            { btnScalePosition, () => new ElementUi { Label = ARES.btnScalePosition_Middle, Image = RES.ScaleFromMiddle } },
            { btnExtendSameLeft, () => new ElementUi { Label = ARES.btnExtendSameLeft, Image = RES.ExtendSameLeft } },
            { btnExtendSameRight, () => new ElementUi { Label = ARES.btnExtendSameRight, Image = RES.ExtendSameRight } },
            { btnExtendSameTop, () => new ElementUi { Label = ARES.btnExtendSameTop, Image = RES.ExtendSameTop } },
            { btnExtendSameBottom, () => new ElementUi { Label = ARES.btnExtendSameBottom, Image = RES.ExtendSameBottom } },
            { btnSnapLeft, () => new ElementUi { Label = ARES.btnSnapLeft, Image = RES.SnapToLeft } },
            { btnSnapRight, () => new ElementUi { Label = ARES.btnSnapRight, Image = RES.SnapToRight } },
            { btnSnapTop, () => new ElementUi { Label = ARES.btnSnapTop, Image = RES.SnapToTop } },
            { btnSnapBottom, () => new ElementUi { Label = ARES.btnSnapBottom, Image = RES.SnapToBottom } },
            { btnMoveForward, () => new ElementUi { Label = ARES.btnMoveForward, Image = RES.ObjectBringForward } },
            { btnMoveFront, () => new ElementUi { Label = ARES.btnMoveFront, Image = RES.ObjectBringToFront } },
            { btnMoveBackward, () => new ElementUi { Label = ARES.btnMoveBackward, Image = RES.ObjectSendBackward } },
            { btnMoveBack, () => new ElementUi { Label = ARES.btnMoveBack, Image = RES.ObjectSendToBack } },
            { btnRotateRight90, () => new ElementUi { Label = ARES.btnRotateRight90, Image = RES.ObjectRotateRight90 } },
            { btnRotateLeft90, () => new ElementUi { Label = ARES.btnRotateLeft90, Image = RES.ObjectRotateLeft90 } },
            { btnFlipVertical, () => new ElementUi { Label = ARES.btnFlipVertical, Image = RES.ObjectFlipVertical } },
            { btnFlipHorizontal, () => new ElementUi { Label = ARES.btnFlipHorizontal, Image = RES.ObjectFlipHorizontal } },
            { btnGroup, () => new ElementUi { Label = ARES.btnGroup, Image = RES.ObjectsGroup } },
            { btnUngroup, () => new ElementUi { Label = ARES.btnUngroup, Image = RES.ObjectsUngroup } },
            // grpTextbox
            { grpTextbox, () => new ElementUi { Label = ARES.grpTextbox, Image = RES.TextboxSetting } },
            { btnAutofitOff, () => new ElementUi { Label = ARES.btnAutofitOff, Image = RES.TextboxAutofitOff } },
            { btnAutofitText, () => new ElementUi { Label = ARES.btnAutofitText, Image = RES.TextboxAutofitText } },
            { btnAutoResize, () => new ElementUi { Label = ARES.btnAutoResize, Image = RES.TextboxAutoResize } },
            { btnWrapText, () => new ElementUi { Label = ARES.btnWrapText, Image = RES.TextboxWrapText_32 } },
            { edtMarginLeft, () => new ElementUi { Label = ARES.edtMarginLeft } },
            { edtMarginRight, () => new ElementUi { Label = ARES.edtMarginRight } },
            { edtMarginTop, () => new ElementUi { Label = ARES.edtMarginTop } },
            { edtMarginBottom, () => new ElementUi { Label = ARES.edtMarginBottom } },
            { btnResetMarginHorizontal, () => new ElementUi { Label = ARES.btnResetMarginHorizontal, Image = RES.TextboxResetMargin } },
            { btnResetMarginVertical, () => new ElementUi { Label = ARES.btnResetMarginVertical, Image = RES.TextboxResetMargin } },
            // grpShapeSizeAndPosition
            { grpShapeSizeAndPosition, () => new ElementUi { Label = ARES.grpShapeSizeAndPosition, Image = RES.SizeAndPosition } },
            { mnuShapeArrangement, () => new ElementUi { Label = ARES.mnuShapeArrangement, Image = RES.ObjectArrangement_32 } },
            { btnShapeScalePosition, () => new ElementUi { Label = ARES.btnScalePosition_Middle, Image = RES.ScaleFromMiddle } },
            { btnLockShapeAspectRatio, () => new ElementUi { Label = ARES.btnLockShapeAspectRatio, Image = RES.ObjectLockAspectRatio } },
            { btnCopyShapeSize, () => new ElementUi { Label = ARES.btnCopyShapeSize, Image = RES.Copy } },
            { btnPasteShapeSize, () => new ElementUi { Label = ARES.btnPasteShapeSize, Image = RES.Paste } },
            { edtShapePositionX, () => new ElementUi { Label = ARES.edtShapePositionX } },
            { edtShapePositionY, () => new ElementUi { Label = ARES.edtShapePositionY } },
            { btnCopyShapePosition, () => new ElementUi { Label = ARES.btnCopyShapePosition, Image = RES.Copy } },
            { btnPasteShapePosition, () => new ElementUi { Label = ARES.btnPasteShapePosition, Image = RES.Paste } },
            // grpReplacePicture
            { grpReplacePicture, () => new ElementUi { Label = ARES.grpReplacePicture, Image = RES.PictureChangeFromClipboard } },
            { btnReplaceWithClipboard, () => new ElementUi { Label = ARES.btnReplaceWithClipboard, Image = RES.PictureChangeFromClipboard_32 } },
            { btnReplaceWithFile, () => new ElementUi { Label = ARES.btnReplaceWithFile, Image = RES.PictureChange } },
            { cbxReserveOriginalSize, () => new ElementUi { Label = ARES.cbxReserveOriginalSize } },
            { cbxReplaceToMiddle, () => new ElementUi { Label = ARES.cbxReplaceToMiddle } },
            // grpPictureSizeAndPosition
            { grpPictureSizeAndPosition, () => new ElementUi { Label = ARES.grpPictureSizeAndPosition, Image = RES.SizeAndPosition } },
            { mnuPictureArrangement, () => new ElementUi { Label = ARES.mnuPictureArrangement, Image = RES.ObjectArrangement_32 } },
            { btnPictureScalePosition, () => new ElementUi { Label = ARES.btnScalePosition_Middle, Image = RES.ScaleFromMiddle } },
            { btnResetPictureSize, () => new ElementUi { Label = ARES.btnResetPictureSize, Image = RES.PictureResetSize_32 } },
            { btnLockPictureAspectRatio, () => new ElementUi { Label = ARES.btnLockPictureAspectRatio, Image = RES.ObjectLockAspectRatio } },
            { btnCopyPictureSize, () => new ElementUi { Label = ARES.btnCopyPictureSize, Image = RES.Copy } },
            { btnPastePictureSize, () => new ElementUi { Label = ARES.btnPastePictureSize, Image = RES.Paste } },
            { edtPicturePositionX, () => new ElementUi { Label = ARES.edtPicturePositionX } },
            { edtPicturePositionY, () => new ElementUi { Label = ARES.edtPicturePositionY } },
            { btnCopyPicturePosition, () => new ElementUi { Label = ARES.btnCopyPicturePosition, Image = RES.Copy } },
            { btnPastePicturePosition, () => new ElementUi { Label = ARES.btnPastePicturePosition, Image = RES.Paste } },
            // mnuArrangement
            { mnuArrangement_sepAlignmentAndResizing, () => new ElementUi { Label = ARES.mnuArrangement_sepAlignmentAndResizing } },
            { mnuArrangement_mnuAlignment, () => new ElementUi { Label = ARES.mnuArrangement_mnuAlignment, Image = RES.ObjectArrangement } },
            { mnuArrangement_mnuResizing, () => new ElementUi { Label = ARES.mnuArrangement_mnuResizing, Image = RES.ScaleSameWidth } },
            { mnuArrangement_mnuSnapping, () => new ElementUi { Label = ARES.mnuArrangement_mnuSnapping, Image = RES.SnapToLeft } },
            { mnuArrangement_mnuRotation, () => new ElementUi { Label = ARES.mnuArrangement_mnuRotation, Image = RES.ObjectRotateRight90 } },
            { mnuArrangement_sepLayerOrderAndGrouping, () => new ElementUi { Label = ARES.mnuArrangement_sepLayerOrderAndGrouping } },
            { mnuArrangement_mnuLayerOrder, () => new ElementUi { Label = ARES.mnuArrangement_mnuLayerOrder, Image = RES.ObjectSendToBack } },
            { mnuArrangement_mnuGrouping, () => new ElementUi { Label = ARES.mnuArrangement_mnuGrouping, Image = RES.ObjectsGroup } },
            { mnuArrangement_sepObjectsInSlide, () => new ElementUi { Label = ARES.mnuArrangement_sepObjectsInSlide } },
        };

        public string GetLabel(Office.IRibbonControl ribbonControl) {
            _elementLabels.TryGetValue(ribbonControl.Id, out var eui);
            return eui?.Invoke().Label ?? "<Unknown>";
        }

        public System.Drawing.Image GetImage(Office.IRibbonControl ribbonControl) {
            _elementLabels.TryGetValue(ribbonControl.Id, out var eui);
            return eui?.Invoke().Image;
        }

        #endregion

    }

}
