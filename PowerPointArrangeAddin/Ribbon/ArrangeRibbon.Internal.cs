﻿using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using PowerPointArrangeAddin.Helper;

#nullable enable

namespace PowerPointArrangeAddin.Ribbon {

    using RL = ArrangeRibbonResources;
    using RIM = Icon.MaterialIconResources;
    using RIF = Icon.FlatIconResources;

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

        #region Ribbon UI Callbacks (Sugar)

        private class FakeRibbonControl : Office.IRibbonControl {
            public FakeRibbonControl(string id) {
                Id = id;
            }

            public string Id { get; }
            public object Context => 0;
            public string Tag => "";
        }

        public string GetLabel(string ribbonControlId) {
            return GetLabel(new FakeRibbonControl(ribbonControlId));
        }

        public System.Drawing.Image? GetImage(string ribbonControlId) {
            return GetImage(new FakeRibbonControl(ribbonControlId));
        }

        public string GetKeytip(string ribbonControlId) {
            return GetKeytip(new FakeRibbonControl(ribbonControlId));
        }

        #endregion

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
        private const string btnRotateRight90 = "btnRotateRight90";
        private const string btnRotateLeft90 = "btnRotateLeft90";
        private const string btnFlipVertical = "btnFlipVertical";
        private const string btnFlipHorizontal = "btnFlipHorizontal";
        private const string btnMoveFront = "btnMoveFront";
        private const string btnMoveBack = "btnMoveBack";
        private const string btnMoveForward = "btnMoveForward";
        private const string btnMoveBackward = "btnMoveBackward";
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
        private const string sepRotate = "sepRotate";
        private const string bgpMoveLayers = "bgpMoveLayers";
        private const string bgpRotate = "bgpRotate";
        private const string bgpGroupObjects = "bgpGroupObjects";
        private const string sepArrangement = "sepArrangement";
        // tabArrangement
        private const string tabArrangement = "tabArrangement";
        // grpAddInSetting
        private const string grpAddInSetting = "grpAddInSetting";
        private const string btnAddInCheckUpdate = "btnAddInCheckUpdate";
        private const string btnAddInHomepage = "btnAddInHomepage";
        private const string btnAddInFeedback = "btnAddInFeedback";
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
            public UiElement(string? label = null, string? imageName = null, string? keytip = null) {
                Label = label;
                ImageName = imageName;
                Keytip = keytip;
            }

            public string? Label { get; }
            public string? ImageName { get; }
            public string? Keytip { get; }

            public System.Drawing.Image? Image { get; private set; }

            public UiElement ApplyNameToImage() {
                if (ImageName != null) {
                    var (resourceManager, culture) = Misc.AddInSetting.Instance.IconStyle == Misc.AddInIconStyle.Office2010
                        ? (RIM.ResourceManager, RIM.Culture) // material
                        : (RIF.ResourceManager, RIF.Culture); // flat
                    Image = (System.Drawing.Image?) resourceManager.GetObject(ImageName, culture);
                }
                return this;
            }
        }

        private Dictionary<string, UiElement> _ribbonUiElements; // id -> ui
        private Dictionary<string, Dictionary<string, UiElement>> _specialRibbonUiElements; // group -> id -> ui

        private (Dictionary<string, UiElement>, Dictionary<string, Dictionary<string, UiElement>>) GenerateNewUiElements() {
            var map = new Dictionary<string, UiElement>();
            var specialMap = new Dictionary<string, Dictionary<string, UiElement>>();

            void Register(string id, UiElement ui) {
                map[id] = ui.ApplyNameToImage();
            }

            void RegisterS(string group, string id, UiElement ui) {
                if (!specialMap.TryGetValue(group, out var m)) {
                    specialMap[group] = new Dictionary<string, UiElement>();
                    m = specialMap[group];
                }
                m[id] = ui.ApplyNameToImage();
            }

            // grpWordArt
            Register(grpWordArt, new UiElement(RL.grpWordArt, nameof(RIM.TextEffectsMenu)));
            // grpArrange
            Register(grpArrange, new UiElement(RL.grpArrange, nameof(RIM.ObjectArrangement)));
            Register(btnAlignLeft, new UiElement(RL.btnAlignLeft, nameof(RIM.ObjectsAlignLeft), "DL"));
            Register(btnAlignCenter, new UiElement(RL.btnAlignCenter, nameof(RIM.ObjectsAlignCenterHorizontal), "DC"));
            Register(btnAlignRight, new UiElement(RL.btnAlignRight, nameof(RIM.ObjectsAlignRight), "DR"));
            Register(btnAlignTop, new UiElement(RL.btnAlignTop, nameof(RIM.ObjectsAlignTop), "DT"));
            Register(btnAlignMiddle, new UiElement(RL.btnAlignMiddle, nameof(RIM.ObjectsAlignMiddleVertical), "DM"));
            Register(btnAlignBottom, new UiElement(RL.btnAlignBottom, nameof(RIM.ObjectsAlignBottom), "DB"));
            Register(btnDistributeHorizontal, new UiElement(RL.btnDistributeHorizontal, nameof(RIM.AlignDistributeHorizontally), "DH"));
            Register(btnDistributeVertical, new UiElement(RL.btnDistributeVertical, nameof(RIM.AlignDistributeVertically), "DV"));
            Register(btnAlignRelative, new UiElement(RL.btnAlignRelative_ToObjects, nameof(RIM.AlignRelativeToObjects), "DA"));
            Register(btnScaleSameWidth, new UiElement(RL.btnScaleSameWidth, nameof(RIM.ScaleSameWidth), "PW"));
            Register(btnScaleSameHeight, new UiElement(RL.btnScaleSameHeight, nameof(RIM.ScaleSameHeight), "PH"));
            Register(btnScaleSameSize, new UiElement(RL.btnScaleSameSize, nameof(RIM.ScaleSameSize), "PS"));
            Register(btnScaleAnchor, new UiElement(RL.btnScaleAnchor_FromTopLeft, nameof(RIM.ScaleFromTopLeft), "PA"));
            Register(btnExtendSameLeft, new UiElement(RL.btnExtendSameLeft, nameof(RIM.ExtendSameLeft), "PL"));
            Register(btnExtendSameRight, new UiElement(RL.btnExtendSameRight, nameof(RIM.ExtendSameRight), "PR"));
            Register(btnExtendSameTop, new UiElement(RL.btnExtendSameTop, nameof(RIM.ExtendSameTop), "PT"));
            Register(btnExtendSameBottom, new UiElement(RL.btnExtendSameBottom, nameof(RIM.ExtendSameBottom), "PB"));
            Register(btnSnapLeft, new UiElement(RL.btnSnapLeft, nameof(RIM.SnapLeftToRight), "PE"));
            Register(btnSnapRight, new UiElement(RL.btnSnapRight, nameof(RIM.SnapRightToLeft), "PI"));
            Register(btnSnapTop, new UiElement(RL.btnSnapTop, nameof(RIM.SnapTopToBottom), "PO"));
            Register(btnSnapBottom, new UiElement(RL.btnSnapBottom, nameof(RIM.SnapBottomToTop), "PM"));
            Register(btnRotateRight90, new UiElement(RL.btnRotateRight90, nameof(RIM.ObjectRotateRight90), "HR"));
            Register(btnRotateLeft90, new UiElement(RL.btnRotateLeft90, nameof(RIM.ObjectRotateLeft90), "HL"));
            Register(btnFlipVertical, new UiElement(RL.btnFlipVertical, nameof(RIM.ObjectFlipVertical), "HV"));
            Register(btnFlipHorizontal, new UiElement(RL.btnFlipHorizontal, nameof(RIM.ObjectFlipHorizontal), "HH"));
            Register(btnMoveFront, new UiElement(RL.btnMoveFront, nameof(RIM.ObjectBringToFront), "HF"));
            Register(btnMoveBack, new UiElement(RL.btnMoveBack, nameof(RIM.ObjectSendToBack), "HB"));
            Register(btnMoveForward, new UiElement(RL.btnMoveForward, nameof(RIM.ObjectBringForward), "HO"));
            Register(btnMoveBackward, new UiElement(RL.btnMoveBackward, nameof(RIM.ObjectSendBackward), "HA"));
            Register(btnGroup, new UiElement(RL.btnGroup, nameof(RIM.ObjectsGroup), "HG"));
            Register(btnUngroup, new UiElement(RL.btnUngroup, nameof(RIM.ObjectsUngroup), "HU"));
            Register(btnGridSetting, new UiElement(RL.btnGridSetting, nameof(RIM.GridSetting), "HD"));
            Register(mnuArrangement, new UiElement(RL.mnuArrangement, nameof(RIM.ObjectArrangement_32), "B"));
            Register(btnAddInSetting, new UiElement(RL.btnAddInSetting, nameof(RIM.AddInOptions), "AS"));
            // tabArrangement
            Register(tabArrangement, new UiElement(RL.tabArrangement, keytip: "M"));
            // grpAddInSetting
            Register(grpAddInSetting, new UiElement(RL.grpAddInSetting, nameof(RIM.AddInOptions)));
            Register(btnAddInCheckUpdate, new UiElement(RL.btnAddInCheckUpdate, nameof(RIM.AddInUpdate), "AU"));
            Register(btnAddInHomepage, new UiElement(RL.btnAddInHomepage, nameof(RIM.AddInHomepage), "AH"));
            Register(btnAddInFeedback, new UiElement(RL.btnAddInFeedback, nameof(RIM.AddInFeedback), "AF"));
            RegisterS(grpAddInSetting, btnAddInSetting, new UiElement(null, nameof(RIM.AddInOptions_32), "T"));
            // grpAlignment
            Register(grpAlignment, new UiElement(RL.grpAlignment, nameof(RIM.ObjectArrangement)));
            Register(lblAlignmentH, new UiElement(RL.lblAlignmentH));
            Register(lblAlignmentV, new UiElement(RL.lblAlignmentV));
            Register(lblDistribute, new UiElement(RL.lblDistribute));
            Register(btnAlignRelative_ToObjects, new UiElement(RL.btnAlignRelative_ToObjects, nameof(RIM.AlignRelativeToObjects), "DO"));
            Register(btnAlignRelative_ToFirstObject, new UiElement(RL.btnAlignRelative_ToFirstObject, nameof(RIM.AlignRelativeToFirstObject), "DF"));
            Register(btnAlignRelative_ToSlide, new UiElement(RL.btnAlignRelative_ToSlide, nameof(RIM.AlignRelativeToSlide), "DS"));
            Register(btnSizeAndPosition, new UiElement(RL.btnSizeAndPosition, nameof(RIM.SizeAndPosition), "DN"));
            // grpSizeAndSnap
            Register(grpSizeAndSnap, new UiElement(RL.grpSizeAndSnap, nameof(RIM.ScaleSameWidth)));
            Register(lblScaleSize, new UiElement(RL.lblScaleSize));
            Register(lblExtendSize, new UiElement(RL.lblExtendSize));
            Register(lblSnapObjects, new UiElement(RL.lblSnapObjects));
            Register(btnScaleAnchor_FromTopLeft, new UiElement(RL.btnScaleAnchor_FromTopLeft, nameof(RIM.ScaleFromTopLeft), "PA"));
            Register(btnScaleAnchor_FromMiddle, new UiElement(RL.btnScaleAnchor_FromMiddle, nameof(RIM.ScaleFromMiddle), "PD"));
            Register(btnScaleAnchor_FromBottomRight, new UiElement(RL.btnScaleAnchor_FromBottomRight, nameof(RIM.ScaleFromBottomRight), "PG"));
            RegisterS(grpSizeAndSnap, btnSizeAndPosition, new UiElement(keytip: "PP"));
            // grpRotateAndFlip
            Register(grpRotateAndFlip, new UiElement(RL.grpRotateAndFlip, nameof(RIM.ObjectRotateRight90)));
            Register(lblRotateObject, new UiElement(RL.lblRotateObject));
            Register(lblFlipObject, new UiElement(RL.lblFlipObject));
            Register(lbl3DRotation, new UiElement(RL.lbl3DRotation));
            Register(edtAngle, new UiElement(RL.edtAngle, keytip: "AE"));
            Register(btnCopyAngle, new UiElement(RL.btnCopyAngle, nameof(RIM.Copy), "AC"));
            Register(btnPasteAngle, new UiElement(RL.btnPasteAngle, nameof(RIM.Paste), "AP"));
            Register(btnResetAngle, new UiElement(RL.btnResetAngle, nameof(RIM.ResetData), "AR"));
            // grpObjectArrange
            Register(grpObjectArrange, new UiElement(RL.grpObjectArrange, nameof(RIM.ObjectSendToBack)));
            Register(lblMoveLayers, new UiElement(RL.lblMoveLayers));
            Register(lblGroupObjects, new UiElement(RL.lblGroupObjects));
            RegisterS(grpObjectArrange, btnGridSetting, new UiElement(null, nameof(RIM.GridSetting_32), "G"));
            RegisterS(grpObjectArrange, btnSizeAndPosition, new UiElement(null, nameof(RIM.SizeAndPosition_32), "N"));
            // grpObjectSize
            Register(grpObjectSize, new UiElement(RL.grpObjectSize, nameof(RIM.ObjectHeight)));
            Register(btnResetSize, new UiElement(RL.btnResetSize, nameof(RIM.PictureResetSize_32), "SR"));
            Register(btnLockAspectRatio, new UiElement(RL.btnLockAspectRatio, nameof(RIM.ObjectLockAspectRatio), "L"));
            RegisterS(grpObjectSize, btnLockAspectRatio, new UiElement(null, nameof(RIM.ObjectLockAspectRatio_32)));
            Register(edtSizeHeight, new UiElement(RL.edtSizeHeight, keytip: "SH"));
            Register(edtSizeWidth, new UiElement(RL.edtSizeWidth, keytip: "SW"));
            Register(btnCopySize, new UiElement(RL.btnCopySize, nameof(RIM.Copy), "SC"));
            Register(btnPasteSize, new UiElement(RL.btnPasteSize, nameof(RIM.Paste), "SV"));
            RegisterS(grpObjectSize, btnSizeAndPosition, new UiElement(keytip: "SN"));
            // grpObjectPosition
            Register(grpObjectPosition, new UiElement(RL.grpObjectPosition, nameof(RIM.ObjectPosition)));
            Register(edtPositionX, new UiElement(RL.edtPositionX, keytip: "PX"));
            Register(edtPositionY, new UiElement(RL.edtPositionY, keytip: "PY"));
            Register(btnCopyPosition, new UiElement(RL.btnCopyPosition, nameof(RIM.Copy), "PC"));
            Register(btnPastePosition, new UiElement(RL.btnPastePosition, nameof(RIM.Paste), "PV"));
            RegisterS(grpObjectPosition, btnSizeAndPosition, new UiElement(keytip: "PN"));
            // grpTextbox
            Register(grpTextbox, new UiElement(RL.grpTextbox, nameof(RIM.TextboxSetting)));
            Register(btnAutofitOff, new UiElement(RL.btnAutofitOff, nameof(RIM.TextboxAutofitOff), "TF"));
            Register(btnAutoShrinkText, new UiElement(RL.btnAutoShrinkText, nameof(RIM.TextboxAutoShrinkText), "TS"));
            Register(btnAutoResizeShape, new UiElement(RL.btnAutoResizeShape, nameof(RIM.TextboxAutoResizeShape), "TR"));
            Register(btnWrapText, new UiElement(RL.btnWrapText, nameof(RIM.TextboxWrapText_32), "TW"));
            Register(lblHorizontalMargin, new UiElement(RL.lblHorizontalMargin));
            Register(btnResetHorizontalMargin, new UiElement(RL.btnResetHorizontalMargin, nameof(RIM.ResetData), "MH"));
            Register(edtMarginLeft, new UiElement(RL.edtMarginLeft, keytip: "ML"));
            Register(edtMarginRight, new UiElement(RL.edtMarginRight, keytip: "MR"));
            Register(lblVerticalMargin, new UiElement(RL.lblVerticalMargin));
            Register(btnResetVerticalMargin, new UiElement(RL.btnResetVerticalMargin, nameof(RIM.ResetData), "MV"));
            Register(edtMarginTop, new UiElement(RL.edtMarginTop, keytip: "MT"));
            Register(edtMarginBottom, new UiElement(RL.edtMarginBottom, keytip: "MB"));
            // grpReplacePicture
            Register(grpReplacePicture, new UiElement(RL.grpReplacePicture, nameof(RIM.PictureChangeFromClipboard)));
            Register(btnReplaceWithClipboard, new UiElement(RL.btnReplaceWithClipboard, nameof(RIM.PictureChangeFromClipboard_32), "TC"));
            Register(btnReplaceWithFile, new UiElement(RL.btnReplaceWithFile, nameof(RIM.PictureChange), "TF"));
            Register(chkReserveOriginalSize, new UiElement(RL.chkReserveOriginalSize, keytip: "TR"));
            Register(chkReplaceToMiddle, new UiElement(RL.chkReplaceToMiddle, keytip: "TM"));
            // grpSizeAndPosition
            Register(grpShapeSizeAndPosition, new UiElement(RL.grpSizeAndPosition, nameof(RIM.SizeAndPosition)));
            Register(grpPictureSizeAndPosition, new UiElement(RL.grpSizeAndPosition, nameof(RIM.SizeAndPosition)));
            Register(grpVideoSizeAndPosition, new UiElement(RL.grpSizeAndPosition, nameof(RIM.SizeAndPosition)));
            Register(grpAudioSizeAndPosition, new UiElement(RL.grpSizeAndPosition, nameof(RIM.SizeAndPosition)));
            Register(grpTableSizeAndPosition, new UiElement(RL.grpSizeAndPosition, nameof(RIM.SizeAndPosition)));
            Register(grpChartSizeAndPosition, new UiElement(RL.grpSizeAndPosition, nameof(RIM.SizeAndPosition)));
            Register(grpSmartartSizeAndPosition, new UiElement(RL.grpSizeAndPosition, nameof(RIM.SizeAndPosition)));
            RegisterS(grpShapeSizeAndPosition, btnSizeAndPosition, new UiElement(keytip: "SN"));
            RegisterS(grpPictureSizeAndPosition, btnSizeAndPosition, new UiElement(keytip: "SN"));
            RegisterS(grpVideoSizeAndPosition, btnSizeAndPosition, new UiElement(keytip: "SN"));
            RegisterS(grpAudioSizeAndPosition, btnSizeAndPosition, new UiElement(keytip: "SN"));
            RegisterS(grpTableSizeAndPosition, btnSizeAndPosition, new UiElement(keytip: "SN"));
            RegisterS(grpChartSizeAndPosition, btnSizeAndPosition, new UiElement(keytip: "SN"));
            RegisterS(grpSmartartSizeAndPosition, btnSizeAndPosition, new UiElement(keytip: "SN"));
            // ===
            RegisterS(grpVideoSizeAndPosition, btnLockAspectRatio, new UiElement(keytip: "SL")); // L
            RegisterS(grpVideoSizeAndPosition, btnScaleAnchor, new UiElement(keytip: "SF")); // PA
            RegisterS(grpVideoSizeAndPosition, edtPositionX, new UiElement(keytip: "SX")); // PX
            RegisterS(grpVideoSizeAndPosition, edtPositionY, new UiElement(keytip: "SY")); // PY
            RegisterS(grpVideoSizeAndPosition, btnCopyPosition, new UiElement(keytip: "SS")); // PC
            RegisterS(grpVideoSizeAndPosition, btnPastePosition, new UiElement(keytip: "ST")); // PV
            // ===
            RegisterS(grpTableSizeAndPosition, mnuArrangement, new UiElement(keytip: "SB")); // B
            RegisterS(grpTableSizeAndPosition, btnLockAspectRatio, new UiElement(keytip: "SL")); // L
            RegisterS(grpTableSizeAndPosition, btnScaleAnchor, new UiElement(keytip: "SF")); // PA
            RegisterS(grpTableSizeAndPosition, edtPositionX, new UiElement(keytip: "SX")); // PX
            RegisterS(grpTableSizeAndPosition, edtPositionY, new UiElement(keytip: "SY")); // PY
            RegisterS(grpTableSizeAndPosition, btnCopyPosition, new UiElement(keytip: "SS")); // PC
            RegisterS(grpTableSizeAndPosition, btnPastePosition, new UiElement(keytip: "ST")); // PV
            // mnuArrangement
            Register(sepAlignmentAndResizing, new UiElement(RL.mnuArrangement_sepAlignmentAndResizing));
            Register(mnuAlignment, new UiElement(RL.mnuArrangement_mnuAlignment, nameof(RIM.ObjectArrangement)));
            Register(mnuResizing, new UiElement(RL.mnuArrangement_mnuResizing, nameof(RIM.ScaleSameWidth)));
            Register(mnuSnapping, new UiElement(RL.mnuArrangement_mnuSnapping, nameof(RIM.SnapLeftToRight)));
            Register(mnuRotation, new UiElement(RL.mnuArrangement_mnuRotation, nameof(RIM.ObjectRotateRight90)));
            Register(sepLayerOrderAndGrouping, new UiElement(RL.mnuArrangement_sepLayerOrderAndGrouping));
            Register(mnuLayerOrder, new UiElement(RL.mnuArrangement_mnuLayerOrder, nameof(RIM.ObjectSendToBack)));
            Register(mnuGrouping, new UiElement(RL.mnuArrangement_mnuGrouping, nameof(RIM.ObjectsGroup)));
            Register(sepObjectsInSlide, new UiElement(RL.mnuArrangement_sepObjectsInSlide));
            Register(sepAddInSetting, new UiElement(RL.mnuArrangement_sepAddInSetting));

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
