using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using PowerPointArrangeAddin.Helper;

#nullable enable

namespace PowerPointArrangeAddin.Ribbon {

    using R1 = ArrangeRibbonResources;
    using R2 = Properties.Resources;

    [ComVisible(true)]
    public partial class ArrangeRibbon : Office.IRibbonExtensibility {

        public string GetCustomUI(string ribbonId) {
            var xml = XmlResourceHelper.GetResourceText(ArrangeRibbonXmlName);
            if (xml == null) {
                return "";
            }

            xml = XmlResourceHelper.ApplyAttributeTemplateForXml(xml);
            xml = XmlResourceHelper.ApplySubtreeTemplateForXml(xml);
            xml = XmlResourceHelper.ApplyMsoKeytipForXml(xml, _msoKeytips);
            xml = XmlResourceHelper.NormalizeControlIdInGroup(xml);
            return xml;
        }

        public string GetMenuContent(Office.IRibbonControl _) {
            var xml = XmlResourceHelper.GetResourceText(ArrangeRibbonMenuXmlName);
            if (xml == null) {
                return "";
            }

            xml = XmlResourceHelper.ApplyAttributeTemplateForXml(xml);
            xml = XmlResourceHelper.ApplySubtreeTemplateForXml(xml);
            xml = XmlResourceHelper.NormalizeControlIdInMenu(xml, mnuArrangement);
            return xml;
        }

        public void UpdateElementUiAndInvalidateRibbon() {
            (_ribbonElementUis, _ribbonElementUiSpecials) = GenerateNewElementUis();
            InvalidateRibbon();
        }

        #region Ribbon UI Callbacks

        private T? GetElementUiField<T>(Office.IRibbonControl ribbonControl, Func<ElementUi, T> getter) {
            if (_ribbonElementUiSpecials.TryGetValue(ribbonControl.Group(), out var m)) {
                if (m.TryGetValue(ribbonControl.Id(), out var eui1) && eui1 != null) {
                    var field = getter(eui1);
                    if (field != null) {
                        return field;
                    }
                }
            }
            _ribbonElementUis.TryGetValue(ribbonControl.Id(), out var eui2);
            return eui2 == null ? default : getter(eui2);
        }

        public string GetLabel(Office.IRibbonControl ribbonControl) {
            return GetElementUiField(ribbonControl, eui => eui.Label) ?? "<Unknown>";
        }

        public System.Drawing.Image? GetImage(Office.IRibbonControl ribbonControl) {
            return GetElementUiField(ribbonControl, eui => eui.Image);
        }

        public string GetKeytip(Office.IRibbonControl ribbonControl) {
            return GetElementUiField(ribbonControl, eui => eui.Keytip) ?? "";
        }

        // Note: The following ui callback methods are defined in "ArrangeRibbon.cs"
        //     - GetEnabled
        //     - GetControlVisible
        //     - GetGroupVisible

        #endregion

        #region Ribbon Element IDs

        // ReSharper disable InconsistentNaming
        // grpWordArt
        private const string grpWordArt = "grpWordArt";
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
        private const string btnAlignRelative = "btnAlignRelative";
        private const string btnScaleSameWidth = "btnScaleSameWidth";
        private const string btnScaleSameHeight = "btnScaleSameHeight";
        private const string btnScaleSameSize = "btnScaleSameSize";
        private const string btnScaleAnchor = "btnScaleAnchor";
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
        private const string mnuArrangement = "mnuArrangement";
        private const string btnAddInSetting = "btnAddInSetting";
        // ===
        private const string bgpAlignLR = "bgpAlignLR";
        private const string bgpAlignTB = "bgpAlignTB";
        private const string bgpDistribute = "bgpDistribute";
        private const string sepScaleSize = "sepScaleSize";
        private const string bgpScaleSize = "bgpScaleSize";
        private const string bgpExtendSize = "bgpExtendSize";
        private const string bgpSnapObjects = "bgpSnapObjects";
        private const string sepMoveLayers = "sepMoveLayers";
        private const string bgpMoveLayers = "bgpMoveLayers";
        private const string bgpRotate = "bgpRotate";
        private const string bgpGroupObjects = "bgpGroupObjects";
        private const string sepArrangement = "sepArrangement";
        // grpTextbox
        private const string grpTextbox = "grpTextbox";
        private const string btnAutofitOff = "btnAutofitOff";
        private const string btnAutoShrinkText = "btnAutoShrinkText";
        private const string btnAutoResizeShape = "btnAutoResizeShape";
        private const string btnWrapText = "btnWrapText";
        private const string lblHorizontalMargin = "lblHorizontalMargin";
        private const string btnResetHorizontalMargin = "btnResetHorizontalMargin";
        private const string edtMarginLeft = "edtMarginLeft";
        private const string edtMarginRight = "edtMarginRight";
        private const string lblVerticalMargin = "lblVerticalMargin";
        private const string btnResetVerticalMargin = "btnResetVerticalMargin";
        private const string edtMarginTop = "edtMarginTop";
        private const string edtMarginBottom = "edtMarginBottom";
        // ===
        private const string sepHorizontalMargin = "sepHorizontalMargin";
        private const string bgpHorizontalMargin = "bgpHorizontalMargin";
        private const string sepVerticalMargin = "sepVerticalMargin";
        private const string bgpVerticalMargin = "bgpVerticalMargin";
        // grpReplacePicture
        private const string grpReplacePicture = "grpReplacePicture";
        private const string btnReplaceWithClipboard = "btnReplaceWithClipboard";
        private const string btnReplaceWithFile = "btnReplaceWithFile";
        private const string chkReserveOriginalSize = "chkReserveOriginalSize";
        private const string chkReplaceToMiddle = "chkReplaceToMiddle";
        // grpSizeAndPosition
        private const string grpShapeSizeAndPosition = "grpShapeSizeAndPosition";
        private const string grpPictureSizeAndPosition = "grpPictureSizeAndPosition";
        private const string grpVideoSizeAndPosition = "grpVideoSizeAndPosition";
        private const string grpAudioSizeAndPosition = "grpAudioSizeAndPosition";
        private const string grpTableSizeAndPosition = "grpTableSizeAndPosition";
        private const string grpChartSizeAndPosition = "grpChartSizeAndPosition";
        private const string grpSmartartSizeAndPosition = "grpSmartartSizeAndPosition";
        private const string btnResetSize = "btnResetSize";
        private const string btnLockAspectRatio = "btnLockAspectRatio";
        private const string btnCopySize = "btnCopySize";
        private const string btnPasteSize = "btnPasteSize";
        private const string edtPositionX = "edtPositionX";
        private const string edtPositionY = "edtPositionY";
        private const string btnCopyPosition = "btnCopyPosition";
        private const string btnPastePosition = "btnPastePosition";
        // ===
        private const string sepResetSize = "sepResetSize";
        private const string bgpCopyAndPasteSize = "bgpCopyAndPasteSize";
        private const string sepPosition = "sepPosition";
        private const string bgpCopyAndPastePosition = "bgpCopyAndPastePosition";
        // mnuArrangement
        private const string sepAlignmentAndResizing = "sepAlignmentAndResizing";
        private const string mnuAlignment = "mnuAlignment";
        private const string btnAlignRelative_ToObjects = "btnAlignRelative_ToObjects";
        private const string btnAlignRelative_ToFirstObject = "btnAlignRelative_ToFirstObject";
        private const string btnAlignRelative_ToSlide = "btnAlignRelative_ToSlide";
        private const string mnuResizing = "mnuResizing";
        private const string btnScaleAnchor_FromTopLeft = "btnScaleAnchor_FromTopLeft";
        private const string btnScaleAnchor_FromMiddle = "btnScaleAnchor_FromMiddle";
        private const string btnScaleAnchor_FromBottomRight = "btnScaleAnchor_FromBottomRight";
        private const string mnuSnapping = "mnuSnapping";
        private const string mnuRotation = "mnuRotation";
        private const string sepLayerOrderAndGrouping = "sepLayerOrderAndGrouping";
        private const string mnuLayerOrder = "mnuLayerOrder";
        private const string mnuGrouping = "mnuGrouping";
        private const string sepObjectsInSlide = "sepObjectsInSlide";
        private const string sepAddInSetting = "sepAddInSetting";
        // ReSharper restore InconsistentNaming

        private string[] _sizeAndPositionGroups = {
            grpShapeSizeAndPosition, grpPictureSizeAndPosition,
            grpVideoSizeAndPosition, grpAudioSizeAndPosition,
            grpTableSizeAndPosition, grpChartSizeAndPosition,
            grpSmartartSizeAndPosition
        };

        #endregion

        #region Ribbon Element UIs

        private class ElementUi {
            public string? Label { get; init; }
            public System.Drawing.Image? Image { get; init; }
            public string? Keytip { get; init; }
        }

        private Dictionary<string, ElementUi> _ribbonElementUis; // id -> ui
        private Dictionary<string, Dictionary<string, ElementUi>> _ribbonElementUiSpecials; // group -> id -> ui

        private (Dictionary<string, ElementUi>, Dictionary<string, Dictionary<string, ElementUi>>) GenerateNewElementUis() {
            var map = new Dictionary<string, ElementUi>();
            var specialMap = new Dictionary<string, Dictionary<string, ElementUi>>();

            void Register1(string id, ElementUi ui) {
                map[id] = ui;
            }

            void Register2(string group, string id, ElementUi ui) {
                if (!specialMap.TryGetValue(group, out var m)) {
                    specialMap[group] = new Dictionary<string, ElementUi>();
                    m = specialMap[group];
                }
                m[id] = ui;
            }

            // grpWordArt
            Register1(grpWordArt, new ElementUi { Label = R1.grpWordArt, Image = R2.TextEffectsMenu });
            // grpArrange
            Register1(grpArrange, new ElementUi { Label = R1.grpArrange, Image = R2.ObjectArrangement });
            Register1(btnAlignLeft, new ElementUi { Label = R1.btnAlignLeft, Image = R2.ObjectsAlignLeft, Keytip = "DL" });
            Register1(btnAlignCenter, new ElementUi { Label = R1.btnAlignCenter, Image = R2.ObjectsAlignCenterHorizontal, Keytip = "DC" });
            Register1(btnAlignRight, new ElementUi { Label = R1.btnAlignRight, Image = R2.ObjectsAlignRight, Keytip = "DR" });
            Register1(btnAlignTop, new ElementUi { Label = R1.btnAlignTop, Image = R2.ObjectsAlignTop, Keytip = "DT" });
            Register1(btnAlignMiddle, new ElementUi { Label = R1.btnAlignMiddle, Image = R2.ObjectsAlignMiddleVertical, Keytip = "DM" });
            Register1(btnAlignBottom, new ElementUi { Label = R1.btnAlignBottom, Image = R2.ObjectsAlignBottom, Keytip = "DB" });
            Register1(btnDistributeHorizontal, new ElementUi { Label = R1.btnDistributeHorizontal, Image = R2.AlignDistributeHorizontally, Keytip = "DH" });
            Register1(btnDistributeVertical, new ElementUi { Label = R1.btnDistributeVertical, Image = R2.AlignDistributeVertically, Keytip = "DV" });
            Register1(btnAlignRelative, new ElementUi { Label = R1.btnAlignRelative_ToObjects, Image = R2.AlignRelativeToObjects, Keytip = "DA" });
            Register1(btnScaleSameWidth, new ElementUi { Label = R1.btnScaleSameWidth, Image = R2.ScaleSameWidth, Keytip = "PW" });
            Register1(btnScaleSameHeight, new ElementUi { Label = R1.btnScaleSameHeight, Image = R2.ScaleSameHeight, Keytip = "PH" });
            Register1(btnScaleSameSize, new ElementUi { Label = R1.btnScaleSameSize, Image = R2.ScaleSameSize, Keytip = "PS" });
            Register1(btnScaleAnchor, new ElementUi { Label = R1.btnScaleAnchor_TopLeft, Image = R2.ScaleFromTopLeft, Keytip = "PA" });
            Register1(btnExtendSameLeft, new ElementUi { Label = R1.btnExtendSameLeft, Image = R2.ExtendSameLeft, Keytip = "PL" });
            Register1(btnExtendSameRight, new ElementUi { Label = R1.btnExtendSameRight, Image = R2.ExtendSameRight, Keytip = "PR" });
            Register1(btnExtendSameTop, new ElementUi { Label = R1.btnExtendSameTop, Image = R2.ExtendSameTop, Keytip = "PT" });
            Register1(btnExtendSameBottom, new ElementUi { Label = R1.btnExtendSameBottom, Image = R2.ExtendSameBottom, Keytip = "PB" });
            Register1(btnSnapLeft, new ElementUi { Label = R1.btnSnapLeft, Image = R2.SnapLeftToRight, Keytip = "PE" });
            Register1(btnSnapRight, new ElementUi { Label = R1.btnSnapRight, Image = R2.SnapRightToLeft, Keytip = "PI" });
            Register1(btnSnapTop, new ElementUi { Label = R1.btnSnapTop, Image = R2.SnapTopToBottom, Keytip = "PO" });
            Register1(btnSnapBottom, new ElementUi { Label = R1.btnSnapBottom, Image = R2.SnapBottomToTop, Keytip = "PM" });
            Register1(btnMoveForward, new ElementUi { Label = R1.btnMoveForward, Image = R2.ObjectBringForward, Keytip = "HF" });
            Register1(btnMoveFront, new ElementUi { Label = R1.btnMoveFront, Image = R2.ObjectBringToFront, Keytip = "HO" });
            Register1(btnMoveBackward, new ElementUi { Label = R1.btnMoveBackward, Image = R2.ObjectSendBackward, Keytip = "HB" });
            Register1(btnMoveBack, new ElementUi { Label = R1.btnMoveBack, Image = R2.ObjectSendToBack, Keytip = "HK" });
            Register1(btnRotateRight90, new ElementUi { Label = R1.btnRotateRight90, Image = R2.ObjectRotateRight90, Keytip = "HR" });
            Register1(btnRotateLeft90, new ElementUi { Label = R1.btnRotateLeft90, Image = R2.ObjectRotateLeft90, Keytip = "HL" });
            Register1(btnFlipVertical, new ElementUi { Label = R1.btnFlipVertical, Image = R2.ObjectFlipVertical, Keytip = "HV" });
            Register1(btnFlipHorizontal, new ElementUi { Label = R1.btnFlipHorizontal, Image = R2.ObjectFlipHorizontal, Keytip = "HH" });
            Register1(btnGroup, new ElementUi { Label = R1.btnGroup, Image = R2.ObjectsGroup, Keytip = "HG" });
            Register1(btnUngroup, new ElementUi { Label = R1.btnUngroup, Image = R2.ObjectsUngroup, Keytip = "HU" });
            Register1(mnuArrangement, new ElementUi { Label = R1.mnuArrangement, Image = R2.ObjectArrangement_32, Keytip = "B" });
            Register1(btnAddInSetting, new ElementUi { Label = R1.btnAddInSetting, Image = R2.AddInOptions, Keytip = "HT" });
            // grpTextbox
            Register1(grpTextbox, new ElementUi { Label = R1.grpTextbox, Image = R2.TextboxSetting });
            Register1(btnAutofitOff, new ElementUi { Label = R1.btnAutofitOff, Image = R2.TextboxAutofitOff, Keytip = "TF" });
            Register1(btnAutoShrinkText, new ElementUi { Label = R1.btnAutoShrinkText, Image = R2.TextboxAutoShrinkText, Keytip = "TS" });
            Register1(btnAutoResizeShape, new ElementUi { Label = R1.btnAutoResizeShape, Image = R2.TextboxAutoResizeShape, Keytip = "TR" });
            Register1(btnWrapText, new ElementUi { Label = R1.btnWrapText, Image = R2.TextboxWrapText_32, Keytip = "TW" });
            Register1(lblHorizontalMargin, new ElementUi { Label = R1.lblHorizontalMargin });
            Register1(btnResetHorizontalMargin, new ElementUi { Label = R1.btnResetHorizontalMargin, Image = R2.TextboxResetMargin, Keytip = "MH" });
            Register1(edtMarginLeft, new ElementUi { Label = R1.edtMarginLeft, Keytip = "ML" });
            Register1(edtMarginRight, new ElementUi { Label = R1.edtMarginRight, Keytip = "MR" });
            Register1(lblVerticalMargin, new ElementUi { Label = R1.lblVerticalMargin });
            Register1(btnResetVerticalMargin, new ElementUi { Label = R1.btnResetVerticalMargin, Image = R2.TextboxResetMargin, Keytip = "MV" });
            Register1(edtMarginTop, new ElementUi { Label = R1.edtMarginTop, Keytip = "MT" });
            Register1(edtMarginBottom, new ElementUi { Label = R1.edtMarginBottom, Keytip = "MB" });
            // grpReplacePicture
            Register1(grpReplacePicture, new ElementUi { Label = R1.grpReplacePicture, Image = R2.PictureChangeFromClipboard });
            Register1(btnReplaceWithClipboard, new ElementUi { Label = R1.btnReplaceWithClipboard, Image = R2.PictureChangeFromClipboard_32, Keytip = "TC" });
            Register1(btnReplaceWithFile, new ElementUi { Label = R1.btnReplaceWithFile, Image = R2.PictureChange, Keytip = "TF" });
            Register1(chkReserveOriginalSize, new ElementUi { Label = R1.chkReserveOriginalSize, Keytip = "TR" });
            Register1(chkReplaceToMiddle, new ElementUi { Label = R1.chkReplaceToMiddle, Keytip = "TM" });
            // grpSizeAndPosition
            Register1(grpShapeSizeAndPosition, new ElementUi { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register1(grpPictureSizeAndPosition, new ElementUi { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register1(grpVideoSizeAndPosition, new ElementUi { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register1(grpAudioSizeAndPosition, new ElementUi { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register1(grpTableSizeAndPosition, new ElementUi { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register1(grpChartSizeAndPosition, new ElementUi { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register1(grpSmartartSizeAndPosition, new ElementUi { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register1(btnResetSize, new ElementUi { Label = R1.btnResetSize, Image = R2.PictureResetSize_32, Keytip = "SR" });
            Register1(btnLockAspectRatio, new ElementUi { Label = R1.btnLockAspectRatio, Image = R2.ObjectLockAspectRatio, Keytip = "L" });
            Register1(btnCopySize, new ElementUi { Label = R1.btnCopySize, Image = R2.Copy, Keytip = "SC" });
            Register1(btnPasteSize, new ElementUi { Label = R1.btnPasteSize, Image = R2.Paste, Keytip = "SP" });
            Register1(edtPositionX, new ElementUi { Label = R1.edtPositionX, Keytip = "PX" });
            Register1(edtPositionY, new ElementUi { Label = R1.edtPositionY, Keytip = "PY" });
            Register1(btnCopyPosition, new ElementUi { Label = R1.btnCopyPosition, Image = R2.Copy, Keytip = "PC" });
            Register1(btnPastePosition, new ElementUi { Label = R1.btnPastePosition, Image = R2.Paste, Keytip = "PP" });
            // ===
            Register2(grpVideoSizeAndPosition, btnLockAspectRatio, new ElementUi { Keytip = "SL" }); // L
            Register2(grpVideoSizeAndPosition, btnScaleAnchor, new ElementUi { Keytip = "SF" }); // PA
            Register2(grpVideoSizeAndPosition, edtPositionX, new ElementUi { Keytip = "SX" }); // PX
            Register2(grpVideoSizeAndPosition, edtPositionY, new ElementUi { Keytip = "SY" }); // PY
            Register2(grpVideoSizeAndPosition, btnCopyPosition, new ElementUi { Keytip = "SS" }); // PC
            Register2(grpVideoSizeAndPosition, btnPastePosition, new ElementUi { Keytip = "ST" }); // PP
            // ===
            Register2(grpTableSizeAndPosition, mnuArrangement, new ElementUi { Keytip = "SB" }); // B
            Register2(grpTableSizeAndPosition, btnLockAspectRatio, new ElementUi { Keytip = "SL" }); // L
            Register2(grpTableSizeAndPosition, btnScaleAnchor, new ElementUi { Keytip = "SF" }); // PA
            Register2(grpTableSizeAndPosition, edtPositionX, new ElementUi { Keytip = "SX" }); // PX
            Register2(grpTableSizeAndPosition, edtPositionY, new ElementUi { Keytip = "SY" }); // PY
            Register2(grpTableSizeAndPosition, btnCopyPosition, new ElementUi { Keytip = "SS" }); // PC
            Register2(grpTableSizeAndPosition, btnPastePosition, new ElementUi { Keytip = "ST" }); // PP
            // mnuArrangement
            Register1(sepAlignmentAndResizing, new ElementUi { Label = R1.mnuArrangement_sepAlignmentAndResizing });
            Register1(mnuAlignment, new ElementUi { Label = R1.mnuArrangement_mnuAlignment, Image = R2.ObjectArrangement });
            Register1(mnuResizing, new ElementUi { Label = R1.mnuArrangement_mnuResizing, Image = R2.ScaleSameWidth });
            Register1(mnuSnapping, new ElementUi { Label = R1.mnuArrangement_mnuSnapping, Image = R2.SnapLeftToRight });
            Register1(mnuRotation, new ElementUi { Label = R1.mnuArrangement_mnuRotation, Image = R2.ObjectRotateRight90 });
            Register1(sepLayerOrderAndGrouping, new ElementUi { Label = R1.mnuArrangement_sepLayerOrderAndGrouping });
            Register1(mnuLayerOrder, new ElementUi { Label = R1.mnuArrangement_mnuLayerOrder, Image = R2.ObjectSendToBack });
            Register1(mnuGrouping, new ElementUi { Label = R1.mnuArrangement_mnuGrouping, Image = R2.ObjectsGroup });
            Register1(sepObjectsInSlide, new ElementUi { Label = R1.mnuArrangement_sepObjectsInSlide });
            Register1(sepAddInSetting, new ElementUi { Label = R1.mnuArrangement_sepAddInSetting });

            return (map, specialMap);
        }

        private readonly Dictionary<string, Dictionary<string, string>> _msoKeytips = new() {
            {
                grpWordArt, new() {
                    { "TextStylesGallery", "AQ" }, { "TextFillColorPicker", "AF" }, { "TextOutlineColorPicker", "AU" },
                    { "TextEffectsMenu", "AE" }, { "WordArtFormatDialog", "AG" }
                }
            },
            { grpArrange, new() { { "ObjectSizeAndPositionDialog", "HS" }, { "SelectionPane", "HP" } } },
            { grpTextbox, new() { { "WordArtFormatDialog", "TG" } } },
            { grpShapeSizeAndPosition, new() { { "ObjectSizeAndPositionDialog", "SN" } } },
            { grpPictureSizeAndPosition, new() { { "ObjectSizeAndPositionDialog", "SN" } } },
            { grpVideoSizeAndPosition, new() { { "ObjectSizeAndPositionDialog", "SN" } } },
            { grpAudioSizeAndPosition, new() { { "ObjectSizeAndPositionDialog", "SN" } } },
            { grpTableSizeAndPosition, new() { { "ObjectSizeAndPositionDialog", "SN" } } },
            { grpChartSizeAndPosition, new() { { "ObjectSizeAndPositionDialog", "SN" } } },
            { grpSmartartSizeAndPosition, new() { { "ObjectSizeAndPositionDialog", "SN" } } }
        };

        #endregion

    }

}
