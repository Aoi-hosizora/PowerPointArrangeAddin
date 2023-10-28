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

            xml = XmlResourceHelper.ApplyTemplateForXml(xml);
            xml = XmlResourceHelper.ApplyMsoKeytipForXml(xml, _msoKeytips);
            xml = XmlResourceHelper.ApplyControlRandomId(xml);
            xml = XmlResourceHelper.NormalizeControlIdInGroup(xml);
            // System.Windows.Forms.Clipboard.SetText(xml);
            return xml;
        }

        public string GetMenuContent(Office.IRibbonControl _) {
            var xml = XmlResourceHelper.GetResourceText(ArrangeRibbonMenuXmlName);
            if (xml == null) {
                return "";
            }

            xml = XmlResourceHelper.ApplyTemplateForXml(xml);
            xml = XmlResourceHelper.ApplyControlRandomId(xml);
            xml = XmlResourceHelper.NormalizeControlIdInMenu(xml, mnuArrangement);
            return xml;
        }

        public void UpdateUiElementsAndInvalidate() {
            (_ribbonUiElements, _specialRibbonUiElements) = GenerateNewUiElements();
            InvalidateRibbon();
        }

        #region Ribbon UI Callbacks

        private T? GetUiElementField<T>(Office.IRibbonControl ribbonControl, Func<UiElement, T> getter) {
            if (_specialRibbonUiElements.TryGetValue(ribbonControl.Group(), out var m)) {
                if (m.TryGetValue(ribbonControl.Id(), out var eui1) && eui1 != null) {
                    var field = getter(eui1);
                    if (field != null) {
                        return field;
                    }
                }
            }
            _ribbonUiElements.TryGetValue(ribbonControl.Id(), out var eui2);
            return eui2 == null ? default : getter(eui2);
        }

        public string GetLabel(Office.IRibbonControl ribbonControl) {
            return GetUiElementField(ribbonControl, eui => eui.Label) ?? "<Unknown>";
        }

        public System.Drawing.Image? GetImage(Office.IRibbonControl ribbonControl) {
            return GetUiElementField(ribbonControl, eui => eui.Image);
        }

        public string GetKeytip(Office.IRibbonControl ribbonControl) {
            return GetUiElementField(ribbonControl, eui => eui.Keytip) ?? "";
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
        private const string btnMoveFront = "btnMoveFront";
        private const string btnMoveBack = "btnMoveBack";
        private const string btnMoveForward = "btnMoveForward";
        private const string btnMoveBackward = "btnMoveBackward";
        private const string btnRotateRight90 = "btnRotateRight90";
        private const string btnRotateLeft90 = "btnRotateLeft90";
        private const string btnFlipVertical = "btnFlipVertical";
        private const string btnFlipHorizontal = "btnFlipHorizontal";
        private const string btnGroup = "btnGroup";
        private const string btnUngroup = "btnUngroup";
        private const string btnGridSetting = "btnGridSetting";
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
        // tabArrangement
        private const string tabArrangement = "tabArrangement";
        // grpAddInSetting
        private const string grpAddInSetting = "grpAddInSetting";
        // grpAlignment
        private const string grpAlignment = "grpAlignment";
        private const string lblAlignmentH = "lblAlignmentH";
        private const string lblAlignmentV = "lblAlignmentV";
        private const string lblDistribute = "lblDistribute";
        private const string btnAlignRelative_ToObjects = "btnAlignRelative_ToObjects";
        private const string btnAlignRelative_ToFirstObject = "btnAlignRelative_ToFirstObject";
        private const string btnAlignRelative_ToSlide = "btnAlignRelative_ToSlide";
        private const string btnSizeAndPosition = "btnSizeAndPosition";
        // ===
        private const string sepAlignRelative = "sepAlignRelative";
        // grpSizeAndSnap
        private const string grpSizeAndSnap = "grpSizeAndSnap";
        private const string lblScaleSize = "lblScaleSize";
        private const string lblExtendSize = "lblExtendSize";
        private const string lblSnapObjects = "lblSnapObjects";
        private const string btnScaleAnchor_FromTopLeft = "btnScaleAnchor_FromTopLeft";
        private const string btnScaleAnchor_FromMiddle = "btnScaleAnchor_FromMiddle";
        private const string btnScaleAnchor_FromBottomRight = "btnScaleAnchor_FromBottomRight";
        // ===
        private const string sepScaleAnchor = "sepScaleAnchor";
        // grpRotateAndFlip
        private const string grpRotateAndFlip = "grpRotateAndFlip";
        private const string lblRotateObject = "lblRotateObject";
        private const string lblFlipObject = "lblFlipObject";
        private const string lbl3DRotation = "lbl3DRotation";
        private const string edtAngle = "edtAngle";
        private const string btnCopyAngle = "btnCopyAngle";
        private const string btnPasteAngle = "btnPasteAngle";
        private const string btnResetAngle = "btnResetAngle";
        // ===
        private const string bgpRotateOnly = "bgpRotateOnly";
        private const string bgpFlipOnly = "bgpFlipOnly";
        private const string bgp3DRotation = "bgp3DRotation";
        private const string sepAngle = "sepAngle";
        private const string bgpCopyAndPasteAngle = "bgpCopyAndPasteAngle";
        // grpObjectArrange
        private const string grpObjectArrange = "grpObjectArrange";
        private const string lblMoveLayers = "lblMoveLayers";
        private const string lblGroupObjects = "lblGroupObjects";
        // ===
        private const string bgpMoveFrontAndBack = "bgpMoveFrontAndBack";
        private const string bgpMoveForwardAndBackward = "bgpMoveForwardAndBackward";
        private const string bgpGroupAndUngroup = "bgpGroupAndUngroup";
        private const string sepGridSettings = "sepGridSettings";
        // grpObjectSize
        private const string grpObjectSize = "grpObjectSize";
        private const string btnResetSize = "btnResetSize";
        private const string btnLockAspectRatio = "btnLockAspectRatio";
        private const string edtSizeHeight = "edtSizeHeight";
        private const string edtSizeWidth = "edtSizeWidth";
        private const string btnCopySize = "btnCopySize";
        private const string btnPasteSize = "btnPasteSize";
        // ===
        private const string sepSize = "sepSize";
        private const string bgpCopyAndPasteSize = "bgpCopyAndPasteSize";
        // grpObjectPosition
        private const string grpObjectPosition = "grpObjectPosition";
        private const string edtPositionX = "edtPositionX";
        private const string edtPositionY = "edtPositionY";
        private const string btnCopyPosition = "btnCopyPosition";
        private const string btnPastePosition = "btnPastePosition";
        // ===
        private const string bgpCopyAndPastePosition = "bgpCopyAndPastePosition";
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
        // ===
        private const string sepResetSize = "sepResetSize";
        private const string sepPosition = "sepPosition";
        // mnuArrangement
        private const string sepAlignmentAndResizing = "sepAlignmentAndResizing";
        private const string mnuAlignment = "mnuAlignment";
        private const string mnuResizing = "mnuResizing";
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

        private class UiElement {
            public string? Label { get; init; }
            public System.Drawing.Image? Image { get; init; }
            public string? Keytip { get; init; }
        }

        private Dictionary<string, UiElement> _ribbonUiElements; // id -> ui
        private Dictionary<string, Dictionary<string, UiElement>> _specialRibbonUiElements; // group -> id -> ui

        private (Dictionary<string, UiElement>, Dictionary<string, Dictionary<string, UiElement>>) GenerateNewUiElements() {
            var map = new Dictionary<string, UiElement>();
            var specialMap = new Dictionary<string, Dictionary<string, UiElement>>();

            void Register(string id, UiElement ui) {
                map[id] = ui;
            }

            void RegisterS(string group, string id, UiElement ui) {
                if (!specialMap.TryGetValue(group, out var m)) {
                    specialMap[group] = new Dictionary<string, UiElement>();
                    m = specialMap[group];
                }
                m[id] = ui;
            }

            // grpWordArt
            Register(grpWordArt, new UiElement { Label = R1.grpWordArt, Image = R2.TextEffectsMenu });
            // grpArrange
            Register(grpArrange, new UiElement { Label = R1.grpArrange, Image = R2.ObjectArrangement });
            Register(btnAlignLeft, new UiElement { Label = R1.btnAlignLeft, Image = R2.ObjectsAlignLeft, Keytip = "DL" });
            Register(btnAlignCenter, new UiElement { Label = R1.btnAlignCenter, Image = R2.ObjectsAlignCenterHorizontal, Keytip = "DC" });
            Register(btnAlignRight, new UiElement { Label = R1.btnAlignRight, Image = R2.ObjectsAlignRight, Keytip = "DR" });
            Register(btnAlignTop, new UiElement { Label = R1.btnAlignTop, Image = R2.ObjectsAlignTop, Keytip = "DT" });
            Register(btnAlignMiddle, new UiElement { Label = R1.btnAlignMiddle, Image = R2.ObjectsAlignMiddleVertical, Keytip = "DM" });
            Register(btnAlignBottom, new UiElement { Label = R1.btnAlignBottom, Image = R2.ObjectsAlignBottom, Keytip = "DB" });
            Register(btnDistributeHorizontal, new UiElement { Label = R1.btnDistributeHorizontal, Image = R2.AlignDistributeHorizontally, Keytip = "DH" });
            Register(btnDistributeVertical, new UiElement { Label = R1.btnDistributeVertical, Image = R2.AlignDistributeVertically, Keytip = "DV" });
            Register(btnAlignRelative, new UiElement { Label = R1.btnAlignRelative_ToObjects, Image = R2.AlignRelativeToObjects, Keytip = "DA" });
            Register(btnScaleSameWidth, new UiElement { Label = R1.btnScaleSameWidth, Image = R2.ScaleSameWidth, Keytip = "PW" });
            Register(btnScaleSameHeight, new UiElement { Label = R1.btnScaleSameHeight, Image = R2.ScaleSameHeight, Keytip = "PH" });
            Register(btnScaleSameSize, new UiElement { Label = R1.btnScaleSameSize, Image = R2.ScaleSameSize, Keytip = "PS" });
            Register(btnScaleAnchor, new UiElement { Label = R1.btnScaleAnchor_FromTopLeft, Image = R2.ScaleFromTopLeft, Keytip = "PA" });
            Register(btnExtendSameLeft, new UiElement { Label = R1.btnExtendSameLeft, Image = R2.ExtendSameLeft, Keytip = "PL" });
            Register(btnExtendSameRight, new UiElement { Label = R1.btnExtendSameRight, Image = R2.ExtendSameRight, Keytip = "PR" });
            Register(btnExtendSameTop, new UiElement { Label = R1.btnExtendSameTop, Image = R2.ExtendSameTop, Keytip = "PT" });
            Register(btnExtendSameBottom, new UiElement { Label = R1.btnExtendSameBottom, Image = R2.ExtendSameBottom, Keytip = "PB" });
            Register(btnSnapLeft, new UiElement { Label = R1.btnSnapLeft, Image = R2.SnapLeftToRight, Keytip = "PE" });
            Register(btnSnapRight, new UiElement { Label = R1.btnSnapRight, Image = R2.SnapRightToLeft, Keytip = "PI" });
            Register(btnSnapTop, new UiElement { Label = R1.btnSnapTop, Image = R2.SnapTopToBottom, Keytip = "PO" });
            Register(btnSnapBottom, new UiElement { Label = R1.btnSnapBottom, Image = R2.SnapBottomToTop, Keytip = "PM" });
            Register(btnMoveFront, new UiElement { Label = R1.btnMoveFront, Image = R2.ObjectBringToFront, Keytip = "HF" });
            Register(btnMoveBack, new UiElement { Label = R1.btnMoveBack, Image = R2.ObjectSendToBack, Keytip = "HB" });
            Register(btnMoveForward, new UiElement { Label = R1.btnMoveForward, Image = R2.ObjectBringForward, Keytip = "HO" });
            Register(btnMoveBackward, new UiElement { Label = R1.btnMoveBackward, Image = R2.ObjectSendBackward, Keytip = "HA" });
            Register(btnRotateRight90, new UiElement { Label = R1.btnRotateRight90, Image = R2.ObjectRotateRight90, Keytip = "HR" });
            Register(btnRotateLeft90, new UiElement { Label = R1.btnRotateLeft90, Image = R2.ObjectRotateLeft90, Keytip = "HL" });
            Register(btnFlipVertical, new UiElement { Label = R1.btnFlipVertical, Image = R2.ObjectFlipVertical, Keytip = "HV" });
            Register(btnFlipHorizontal, new UiElement { Label = R1.btnFlipHorizontal, Image = R2.ObjectFlipHorizontal, Keytip = "HH" });
            Register(btnGroup, new UiElement { Label = R1.btnGroup, Image = R2.ObjectsGroup, Keytip = "HG" });
            Register(btnUngroup, new UiElement { Label = R1.btnUngroup, Image = R2.ObjectsUngroup, Keytip = "HU" });
            Register(btnGridSetting, new UiElement { Label = R1.btnGridSetting, Image = R2.GridSetting, Keytip = "HD" });
            Register(mnuArrangement, new UiElement { Label = R1.mnuArrangement, Image = R2.ObjectArrangement_32, Keytip = "B" });
            Register(btnAddInSetting, new UiElement { Label = R1.btnAddInSetting, Image = R2.AddInOptions, Keytip = "AS" });
            // tabArrangement
            Register(tabArrangement, new UiElement { Label = R1.tabArrangement, Keytip = "M" });
            // grpAddInSetting
            Register(grpAddInSetting, new UiElement { Label = R1.grpAddInSetting, Image = R2.AddInOptions });
            RegisterS(grpAddInSetting, btnAddInSetting, new UiElement { Image = R2.AddInOptions_32, Keytip = "T" });
            // grpAlignment
            Register(grpAlignment, new UiElement { Label = R1.grpAlignment, Image = R2.ObjectArrangement });
            Register(lblAlignmentH, new UiElement { Label = R1.lblAlignmentH });
            Register(lblAlignmentV, new UiElement { Label = R1.lblAlignmentV });
            Register(lblDistribute, new UiElement { Label = R1.lblDistribute });
            Register(btnAlignRelative_ToObjects, new UiElement { Label = R1.btnAlignRelative_ToObjects, Image = R2.AlignRelativeToObjects, Keytip = "DO" });
            Register(btnAlignRelative_ToFirstObject, new UiElement { Label = R1.btnAlignRelative_ToFirstObject, Image = R2.AlignRelativeToFirstObject, Keytip = "DF" });
            Register(btnAlignRelative_ToSlide, new UiElement { Label = R1.btnAlignRelative_ToSlide, Image = R2.AlignRelativeToSlide, Keytip = "DS" });
            Register(btnSizeAndPosition, new UiElement { Label = R1.btnSizeAndPosition, Image = R2.SizeAndPosition, Keytip = "DN" });
            // grpSizeAndSnap
            Register(grpSizeAndSnap, new UiElement { Label = R1.grpSizeAndSnap, Image = R2.ScaleSameWidth });
            Register(lblScaleSize, new UiElement { Label = R1.lblScaleSize });
            Register(lblExtendSize, new UiElement { Label = R1.lblExtendSize });
            Register(lblSnapObjects, new UiElement { Label = R1.lblSnapObjects });
            Register(btnScaleAnchor_FromTopLeft, new UiElement { Label = R1.btnScaleAnchor_FromTopLeft, Image = R2.ScaleFromTopLeft, Keytip = "PA" });
            Register(btnScaleAnchor_FromMiddle, new UiElement { Label = R1.btnScaleAnchor_FromMiddle, Image = R2.ScaleFromMiddle, Keytip = "PD" });
            Register(btnScaleAnchor_FromBottomRight, new UiElement { Label = R1.btnScaleAnchor_FromBottomRight, Image = R2.ScaleFromBottomRight, Keytip = "PG" });
            RegisterS(grpSizeAndSnap, btnSizeAndPosition, new UiElement { Keytip = "PP" });
            // grpRotateAndFlip
            Register(grpRotateAndFlip, new UiElement { Label = R1.grpRotateAndFlip, Image = R2.ObjectRotateRight90 });
            Register(lblRotateObject, new UiElement { Label = R1.lblRotateObject });
            Register(lblFlipObject, new UiElement { Label = R1.lblFlipObject });
            Register(lbl3DRotation, new UiElement { Label = R1.lbl3DRotation });
            Register(edtAngle, new UiElement { Label = R1.edtAngle, Keytip = "AE" });
            Register(btnCopyAngle, new UiElement { Label = R1.btnCopyAngle, Image = R2.Copy, Keytip = "AC" });
            Register(btnPasteAngle, new UiElement { Label = R1.btnPasteAngle, Image = R2.Paste, Keytip = "AP" });
            Register(btnResetAngle, new UiElement { Label = R1.btnResetAngle, Image = R2.TextboxResetMargin, Keytip = "AR" });
            // grpObjectArrange
            Register(grpObjectArrange, new UiElement { Label = R1.grpObjectArrange, Image = R2.ObjectSendToBack });
            Register(lblMoveLayers, new UiElement { Label = R1.lblMoveLayers });
            Register(lblGroupObjects, new UiElement { Label = R1.lblGroupObjects });
            RegisterS(grpObjectArrange, btnGridSetting, new UiElement { Image = R2.GridSetting_32, Keytip = "G" });
            RegisterS(grpObjectArrange, btnSizeAndPosition, new UiElement { Image = R2.SizeAndPosition_32, Keytip = "N" });
            // grpObjectSize
            Register(grpObjectSize, new UiElement { Label = R1.grpObjectSize, Image = R2.ObjectHeight });
            Register(btnResetSize, new UiElement { Label = R1.btnResetSize, Image = R2.PictureResetSize_32, Keytip = "SR" });
            Register(btnLockAspectRatio, new UiElement { Label = R1.btnLockAspectRatio, Image = R2.ObjectLockAspectRatio, Keytip = "L" });
            RegisterS(grpObjectSize, btnLockAspectRatio, new UiElement { Image = R2.ObjectLockAspectRatio_32 });
            Register(edtSizeHeight, new UiElement { Label = R1.edtSizeHeight, Keytip = "SH" });
            Register(edtSizeWidth, new UiElement { Label = R1.edtSizeWidth, Keytip = "SW" });
            Register(btnCopySize, new UiElement { Label = R1.btnCopySize, Image = R2.Copy, Keytip = "SC" });
            Register(btnPasteSize, new UiElement { Label = R1.btnPasteSize, Image = R2.Paste, Keytip = "SV" });
            RegisterS(grpObjectSize, btnSizeAndPosition, new UiElement { Keytip = "SN" });
            // grpObjectPosition
            Register(grpObjectPosition, new UiElement { Label = R1.grpObjectPosition, Image = R2.ObjectPosition });
            Register(edtPositionX, new UiElement { Label = R1.edtPositionX, Keytip = "PX" });
            Register(edtPositionY, new UiElement { Label = R1.edtPositionY, Keytip = "PY" });
            Register(btnCopyPosition, new UiElement { Label = R1.btnCopyPosition, Image = R2.Copy, Keytip = "PC" });
            Register(btnPastePosition, new UiElement { Label = R1.btnPastePosition, Image = R2.Paste, Keytip = "PV" });
            RegisterS(grpObjectPosition, btnSizeAndPosition, new UiElement { Keytip = "PN" });
            // grpTextbox
            Register(grpTextbox, new UiElement { Label = R1.grpTextbox, Image = R2.TextboxSetting });
            Register(btnAutofitOff, new UiElement { Label = R1.btnAutofitOff, Image = R2.TextboxAutofitOff, Keytip = "TF" });
            Register(btnAutoShrinkText, new UiElement { Label = R1.btnAutoShrinkText, Image = R2.TextboxAutoShrinkText, Keytip = "TS" });
            Register(btnAutoResizeShape, new UiElement { Label = R1.btnAutoResizeShape, Image = R2.TextboxAutoResizeShape, Keytip = "TR" });
            Register(btnWrapText, new UiElement { Label = R1.btnWrapText, Image = R2.TextboxWrapText_32, Keytip = "TW" });
            Register(lblHorizontalMargin, new UiElement { Label = R1.lblHorizontalMargin });
            Register(btnResetHorizontalMargin, new UiElement { Label = R1.btnResetHorizontalMargin, Image = R2.TextboxResetMargin, Keytip = "MH" });
            Register(edtMarginLeft, new UiElement { Label = R1.edtMarginLeft, Keytip = "ML" });
            Register(edtMarginRight, new UiElement { Label = R1.edtMarginRight, Keytip = "MR" });
            Register(lblVerticalMargin, new UiElement { Label = R1.lblVerticalMargin });
            Register(btnResetVerticalMargin, new UiElement { Label = R1.btnResetVerticalMargin, Image = R2.TextboxResetMargin, Keytip = "MV" });
            Register(edtMarginTop, new UiElement { Label = R1.edtMarginTop, Keytip = "MT" });
            Register(edtMarginBottom, new UiElement { Label = R1.edtMarginBottom, Keytip = "MB" });
            // grpReplacePicture
            Register(grpReplacePicture, new UiElement { Label = R1.grpReplacePicture, Image = R2.PictureChangeFromClipboard });
            Register(btnReplaceWithClipboard, new UiElement { Label = R1.btnReplaceWithClipboard, Image = R2.PictureChangeFromClipboard_32, Keytip = "TC" });
            Register(btnReplaceWithFile, new UiElement { Label = R1.btnReplaceWithFile, Image = R2.PictureChange, Keytip = "TF" });
            Register(chkReserveOriginalSize, new UiElement { Label = R1.chkReserveOriginalSize, Keytip = "TR" });
            Register(chkReplaceToMiddle, new UiElement { Label = R1.chkReplaceToMiddle, Keytip = "TM" });
            // grpSizeAndPosition
            Register(grpShapeSizeAndPosition, new UiElement { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register(grpPictureSizeAndPosition, new UiElement { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register(grpVideoSizeAndPosition, new UiElement { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register(grpAudioSizeAndPosition, new UiElement { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register(grpTableSizeAndPosition, new UiElement { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register(grpChartSizeAndPosition, new UiElement { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            Register(grpSmartartSizeAndPosition, new UiElement { Label = R1.grpSizeAndPosition, Image = R2.SizeAndPosition });
            RegisterS(grpShapeSizeAndPosition, btnSizeAndPosition, new UiElement { Keytip = "SN" });
            RegisterS(grpPictureSizeAndPosition, btnSizeAndPosition, new UiElement { Keytip = "SN" });
            RegisterS(grpVideoSizeAndPosition, btnSizeAndPosition, new UiElement { Keytip = "SN" });
            RegisterS(grpAudioSizeAndPosition, btnSizeAndPosition, new UiElement { Keytip = "SN" });
            RegisterS(grpTableSizeAndPosition, btnSizeAndPosition, new UiElement { Keytip = "SN" });
            RegisterS(grpChartSizeAndPosition, btnSizeAndPosition, new UiElement { Keytip = "SN" });
            RegisterS(grpSmartartSizeAndPosition, btnSizeAndPosition, new UiElement { Keytip = "SN" });
            // ===
            RegisterS(grpVideoSizeAndPosition, btnLockAspectRatio, new UiElement { Keytip = "SL" }); // L
            RegisterS(grpVideoSizeAndPosition, btnScaleAnchor, new UiElement { Keytip = "SF" }); // PA
            RegisterS(grpVideoSizeAndPosition, edtPositionX, new UiElement { Keytip = "SX" }); // PX
            RegisterS(grpVideoSizeAndPosition, edtPositionY, new UiElement { Keytip = "SY" }); // PY
            RegisterS(grpVideoSizeAndPosition, btnCopyPosition, new UiElement { Keytip = "SS" }); // PC
            RegisterS(grpVideoSizeAndPosition, btnPastePosition, new UiElement { Keytip = "ST" }); // PV
            // ===
            RegisterS(grpTableSizeAndPosition, mnuArrangement, new UiElement { Keytip = "SB" }); // B
            RegisterS(grpTableSizeAndPosition, btnLockAspectRatio, new UiElement { Keytip = "SL" }); // L
            RegisterS(grpTableSizeAndPosition, btnScaleAnchor, new UiElement { Keytip = "SF" }); // PA
            RegisterS(grpTableSizeAndPosition, edtPositionX, new UiElement { Keytip = "SX" }); // PX
            RegisterS(grpTableSizeAndPosition, edtPositionY, new UiElement { Keytip = "SY" }); // PY
            RegisterS(grpTableSizeAndPosition, btnCopyPosition, new UiElement { Keytip = "SS" }); // PC
            RegisterS(grpTableSizeAndPosition, btnPastePosition, new UiElement { Keytip = "ST" }); // PV
            // mnuArrangement
            Register(sepAlignmentAndResizing, new UiElement { Label = R1.mnuArrangement_sepAlignmentAndResizing });
            Register(mnuAlignment, new UiElement { Label = R1.mnuArrangement_mnuAlignment, Image = R2.ObjectArrangement });
            Register(mnuResizing, new UiElement { Label = R1.mnuArrangement_mnuResizing, Image = R2.ScaleSameWidth });
            Register(mnuSnapping, new UiElement { Label = R1.mnuArrangement_mnuSnapping, Image = R2.SnapLeftToRight });
            Register(mnuRotation, new UiElement { Label = R1.mnuArrangement_mnuRotation, Image = R2.ObjectRotateRight90 });
            Register(sepLayerOrderAndGrouping, new UiElement { Label = R1.mnuArrangement_sepLayerOrderAndGrouping });
            Register(mnuLayerOrder, new UiElement { Label = R1.mnuArrangement_mnuLayerOrder, Image = R2.ObjectSendToBack });
            Register(mnuGrouping, new UiElement { Label = R1.mnuArrangement_mnuGrouping, Image = R2.ObjectsGroup });
            Register(sepObjectsInSlide, new UiElement { Label = R1.mnuArrangement_sepObjectsInSlide });
            Register(sepAddInSetting, new UiElement { Label = R1.mnuArrangement_sepAddInSetting });

            return (map, specialMap);
        }

        private readonly Dictionary<string, Dictionary<string, string>> _msoKeytips = new() {
            {
                grpWordArt, new() {
                    { "TextStylesGallery", "AQ" }, { "TextFillColorPicker", "AF" }, { "TextOutlineColorPicker", "AU" },
                    { "TextEffectsMenu", "AE" }, { "WordArtFormatDialog", "AG" }
                }
            },
            { grpArrange, new() { { "SelectionPane", "HS" } } },
            { grpRotateAndFlip, new() { { "_3DRotationGallery", "AD" }, { "ObjectRotationOptionsDialog", "AM" } } },
            { grpObjectArrange, new() { { "SelectionPane", "HS" } } },
            { grpObjectSize, new() { { "PictureCropTools", "SP" } } },
            { grpTextbox, new() { { "WordArtFormatDialog", "TG" } } }
        };

        #endregion

    }

}
