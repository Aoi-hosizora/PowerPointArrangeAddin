﻿using System.Collections.Generic;
using System.Runtime.InteropServices;
using ppt_arrange_addin.Helper;
using Office = Microsoft.Office.Core;

#nullable enable

namespace ppt_arrange_addin.Ribbon {

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
            xml = XmlResourceHelper.ApplySubtreeTemplateForXml(xml);
            xml = XmlResourceHelper.ApplyMsoKeytipForXml(xml, _msoKeytips);
            return xml;
        }

        public string GetMenuContent(Office.IRibbonControl _) {
            var xml = XmlResourceHelper.GetResourceText(ArrangeRibbonMenuXmlName);
            if (xml == null) {
                return "";
            }

            xml = XmlResourceHelper.ApplyTemplateForXml(xml);
            return xml;
        }

        public void UpdateElementUiAndInvalidateRibbon() {
            _ribbonElementUis = GenerateNewElementUis();
            InvalidateRibbon();
        }

        #region Ribbon UI Callbacks

        public string GetLabel(Office.IRibbonControl ribbonControl) {
            _ribbonElementUis.TryGetValue(ribbonControl.Id, out var eui);
            return eui?.Label ?? "<Unknown>";
        }

        public System.Drawing.Image? GetImage(Office.IRibbonControl ribbonControl) {
            _ribbonElementUis.TryGetValue(ribbonControl.Id, out var eui);
            return eui?.Image;
        }

        public string GetKeytip(Office.IRibbonControl ribbonControl) {
            _ribbonElementUis.TryGetValue(ribbonControl.Id, out var eui);
            return eui?.Keytip ?? "";
        }

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
        private const string grpArrange_separator1 = "grpArrange_separator1";
        private const string bgpScaleSize = "bgpScaleSize";
        private const string bgpExtendSize = "bgpExtendSize";
        private const string bgpSnapObjects = "bgpSnapObjects";
        private const string grpArrange_separator2 = "grpArrange_separator2";
        private const string bgpMoveLayers = "bgpMoveLayers";
        private const string bgpRotate = "bgpRotate";
        private const string bgpGroupObjects = "bgpGroupObjects";
        private const string grpArrange_separator3 = "grpArrange_separator3";
        // grpTextbox
        private const string grpTextbox = "grpTextbox";
        private const string btnAutofitOff = "btnAutofitOff";
        private const string btnAutoShrinkText = "btnAutoShrinkText";
        private const string btnAutoResizeShape = "btnAutoResizeShape";
        private const string btnWrapText = "btnWrapText";
        private const string edtMarginLeft = "edtMarginLeft";
        private const string edtMarginRight = "edtMarginRight";
        private const string edtMarginTop = "edtMarginTop";
        private const string edtMarginBottom = "edtMarginBottom";
        private const string btnResetHorizontalMargin = "btnResetHorizontalMargin";
        private const string btnResetVerticalMargin = "btnResetVerticalMargin";
        // grpShapeSizeAndPosition
        private const string grpShapeSizeAndPosition = "grpShapeSizeAndPosition";
        private const string mnuShapeArrangement = "mnuShapeArrangement";
        private const string btnLockShapeAspectRatio = "btnLockShapeAspectRatio";
        private const string btnShapeScaleAnchor = "btnShapeScaleAnchor";
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
        private const string chkReserveOriginalSize = "chkReserveOriginalSize";
        private const string chkReplaceToMiddle = "chkReplaceToMiddle";
        // grpPictureSizeAndPosition
        private const string grpPictureSizeAndPosition = "grpPictureSizeAndPosition";
        private const string mnuPictureArrangement = "mnuPictureArrangement";
        private const string btnResetPictureSize = "btnResetPictureSize";
        private const string btnLockPictureAspectRatio = "btnLockPictureAspectRatio";
        private const string btnPictureScaleAnchor = "btnPictureScaleAnchor";
        private const string btnCopyPictureSize = "btnCopyPictureSize";
        private const string btnPastePictureSize = "btnPastePictureSize";
        private const string edtPicturePositionX = "edtPicturePositionX";
        private const string edtPicturePositionY = "edtPicturePositionY";
        private const string btnCopyPicturePosition = "btnCopyPicturePosition";
        private const string btnPastePicturePosition = "btnPastePicturePosition";
        // grpVideoSizeAndPosition
        private const string grpVideoSizeAndPosition = "grpVideoSizeAndPosition";
        private const string mnuVideoArrangement = "mnuVideoArrangement";
        private const string btnResetVideoSize = "btnResetVideoSize";
        private const string btnLockVideoAspectRatio = "btnLockVideoAspectRatio";
        private const string btnVideoScaleAnchor = "btnVideoScaleAnchor";
        private const string btnCopyVideoSize = "btnCopyVideoSize";
        private const string btnPasteVideoSize = "btnPasteVideoSize";
        private const string edtVideoPositionX = "edtVideoPositionX";
        private const string edtVideoPositionY = "edtVideoPositionY";
        private const string btnCopyVideoPosition = "btnCopyVideoPosition";
        private const string btnPasteVideoPosition = "btnPasteVideoPosition";
        // grpAudioSizeAndPosition
        private const string grpAudioSizeAndPosition = "grpAudioSizeAndPosition";
        private const string mnuAudioArrangement = "mnuAudioArrangement";
        private const string btnResetAudioSize = "btnResetAudioSize";
        private const string btnLockAudioAspectRatio = "btnLockAudioAspectRatio";
        private const string btnAudioScaleAnchor = "btnAudioScaleAnchor";
        private const string btnCopyAudioSize = "btnCopyAudioSize";
        private const string btnPasteAudioSize = "btnPasteAudioSize";
        private const string edtAudioPositionX = "edtAudioPositionX";
        private const string edtAudioPositionY = "edtAudioPositionY";
        private const string btnCopyAudioPosition = "btnCopyAudioPosition";
        private const string btnPasteAudioPosition = "btnPasteAudioPosition";
        // grpTableSizeAndPosition
        private const string grpTableSizeAndPosition = "grpTableSizeAndPosition";
        private const string mnuTableArrangement = "mnuTableArrangement";
        private const string btnLockTableAspectRatio = "btnLockTableAspectRatio";
        private const string btnTableScaleAnchor = "btnTableScaleAnchor";
        private const string btnCopyTableSize = "btnCopyTableSize";
        private const string btnPasteTableSize = "btnPasteTableSize";
        private const string edtTablePositionX = "edtTablePositionX";
        private const string edtTablePositionY = "edtTablePositionY";
        private const string btnCopyTablePosition = "btnCopyTablePosition";
        private const string btnPasteTablePosition = "btnPasteTablePosition";
        // grpChartSizeAndPosition
        private const string grpChartSizeAndPosition = "grpChartSizeAndPosition";
        private const string mnuChartArrangement = "mnuChartArrangement";
        private const string btnLockChartAspectRatio = "btnLockChartAspectRatio";
        private const string btnChartScaleAnchor = "btnChartScaleAnchor";
        private const string btnCopyChartSize = "btnCopyChartSize";
        private const string btnPasteChartSize = "btnPasteChartSize";
        private const string edtChartPositionX = "edtChartPositionX";
        private const string edtChartPositionY = "edtChartPositionY";
        private const string btnCopyChartPosition = "btnCopyChartPosition";
        private const string btnPasteChartPosition = "btnPasteChartPosition";
        // grpSmartartSizeAndPosition
        private const string grpSmartartSizeAndPosition = "grpSmartartSizeAndPosition";
        private const string mnuSmartartArrangement = "mnuSmartartArrangement";
        private const string btnLockSmartartAspectRatio = "btnLockSmartartAspectRatio";
        private const string btnSmartartScaleAnchor = "btnSmartartScaleAnchor";
        private const string btnCopySmartartSize = "btnCopySmartartSize";
        private const string btnPasteSmartartSize = "btnPasteSmartartSize";
        private const string edtSmartartPositionX = "edtSmartartPositionX";
        private const string edtSmartartPositionY = "edtSmartartPositionY";
        private const string btnCopySmartartPosition = "btnCopySmartartPosition";
        private const string btnPasteSmartartPosition = "btnPasteSmartartPosition";
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
        private const string mnuArrangement_sepAddInSetting = "mnuArrangement_sepAddInSetting";
        // ReSharper restore InconsistentNaming

        #endregion

        #region Ribbon Element UIs

        private class ElementUi {
            public string? Label { get; init; }
            public System.Drawing.Image? Image { get; init; }
            public string? Keytip { get; init; }
        }

        private Dictionary<string, ElementUi> _ribbonElementUis;

        private Dictionary<string, ElementUi> GenerateNewElementUis() {
            return new Dictionary<string, ElementUi> {
                // grpWordArt
                { grpWordArt, new ElementUi { Label = R1.grpWordArt, Image = R2.TextEffectsMenu } },
                // grpArrange
                { grpArrange, new ElementUi { Label = R1.grpArrange, Image = R2.ObjectArrangement } },
                { btnAlignLeft, new ElementUi { Label = R1.btnAlignLeft, Image = R2.ObjectsAlignLeft, Keytip = "DL" } },
                { btnAlignCenter, new ElementUi { Label = R1.btnAlignCenter, Image = R2.ObjectsAlignCenterHorizontal, Keytip = "DC" } },
                { btnAlignRight, new ElementUi { Label = R1.btnAlignRight, Image = R2.ObjectsAlignRight, Keytip = "DR" } },
                { btnAlignTop, new ElementUi { Label = R1.btnAlignTop, Image = R2.ObjectsAlignTop, Keytip = "DT" } },
                { btnAlignMiddle, new ElementUi { Label = R1.btnAlignMiddle, Image = R2.ObjectsAlignMiddleVertical, Keytip = "DM" } },
                { btnAlignBottom, new ElementUi { Label = R1.btnAlignBottom, Image = R2.ObjectsAlignBottom, Keytip = "DB" } },
                { btnDistributeHorizontal, new ElementUi { Label = R1.btnDistributeHorizontal, Image = R2.AlignDistributeHorizontally, Keytip = "DH" } },
                { btnDistributeVertical, new ElementUi { Label = R1.btnDistributeVertical, Image = R2.AlignDistributeVertically, Keytip = "DV" } },
                { btnScaleSameWidth, new ElementUi { Label = R1.btnScaleSameWidth, Image = R2.ScaleSameWidth, Keytip = "PW" } },
                { btnScaleSameHeight, new ElementUi { Label = R1.btnScaleSameHeight, Image = R2.ScaleSameHeight, Keytip = "PH" } },
                { btnScaleSameSize, new ElementUi { Label = R1.btnScaleSameSize, Image = R2.ScaleSameSize, Keytip = "PS" } },
                { btnScaleAnchor, new ElementUi { Label = R1.btnScaleAnchor_TopLeft, Image = R2.ScaleFromTopLeft, Keytip = "PA" } },
                { btnExtendSameLeft, new ElementUi { Label = R1.btnExtendSameLeft, Image = R2.ExtendSameLeft, Keytip = "PL" } },
                { btnExtendSameRight, new ElementUi { Label = R1.btnExtendSameRight, Image = R2.ExtendSameRight, Keytip = "PR" } },
                { btnExtendSameTop, new ElementUi { Label = R1.btnExtendSameTop, Image = R2.ExtendSameTop, Keytip = "PT" } },
                { btnExtendSameBottom, new ElementUi { Label = R1.btnExtendSameBottom, Image = R2.ExtendSameBottom, Keytip = "PB" } },
                { btnSnapLeft, new ElementUi { Label = R1.btnSnapLeft, Image = R2.SnapLeftToRight, Keytip = "PE" } },
                { btnSnapRight, new ElementUi { Label = R1.btnSnapRight, Image = R2.SnapRightToLeft, Keytip = "PI" } },
                { btnSnapTop, new ElementUi { Label = R1.btnSnapTop, Image = R2.SnapTopToBottom, Keytip = "PO" } },
                { btnSnapBottom, new ElementUi { Label = R1.btnSnapBottom, Image = R2.SnapBottomToTop, Keytip = "PM" } },
                { btnMoveForward, new ElementUi { Label = R1.btnMoveForward, Image = R2.ObjectBringForward, Keytip = "HF" } },
                { btnMoveFront, new ElementUi { Label = R1.btnMoveFront, Image = R2.ObjectBringToFront, Keytip = "HO" } },
                { btnMoveBackward, new ElementUi { Label = R1.btnMoveBackward, Image = R2.ObjectSendBackward, Keytip = "HB" } },
                { btnMoveBack, new ElementUi { Label = R1.btnMoveBack, Image = R2.ObjectSendToBack, Keytip = "HK" } },
                { btnRotateRight90, new ElementUi { Label = R1.btnRotateRight90, Image = R2.ObjectRotateRight90, Keytip = "HR" } },
                { btnRotateLeft90, new ElementUi { Label = R1.btnRotateLeft90, Image = R2.ObjectRotateLeft90, Keytip = "HL" } },
                { btnFlipVertical, new ElementUi { Label = R1.btnFlipVertical, Image = R2.ObjectFlipVertical, Keytip = "HV" } },
                { btnFlipHorizontal, new ElementUi { Label = R1.btnFlipHorizontal, Image = R2.ObjectFlipHorizontal, Keytip = "HH" } },
                { btnGroup, new ElementUi { Label = R1.btnGroup, Image = R2.ObjectsGroup, Keytip = "HG" } },
                { btnUngroup, new ElementUi { Label = R1.btnUngroup, Image = R2.ObjectsUngroup, Keytip = "HU" } },
                { mnuArrangement, new ElementUi { Label = R1.mnuArrangement, Image = R2.ObjectArrangement_32, Keytip = "B" } },
                { btnAddInSetting, new ElementUi { Label = R1.btnAddInSetting, Image = R2.AddInOptions, Keytip = "HT" } },
                // grpTextbox
                { grpTextbox, new ElementUi { Label = R1.grpTextbox, Image = R2.TextboxSetting } },
                { btnAutofitOff, new ElementUi { Label = R1.btnAutofitOff, Image = R2.TextboxAutofitOff, Keytip = "TF" } },
                { btnAutoShrinkText, new ElementUi { Label = R1.btnAutoShrinkText, Image = R2.TextboxAutoShrinkText, Keytip = "TS" } },
                { btnAutoResizeShape, new ElementUi { Label = R1.btnAutoResizeShape, Image = R2.TextboxAutoResizeShape, Keytip = "TR" } },
                { btnWrapText, new ElementUi { Label = R1.btnWrapText, Image = R2.TextboxWrapText_32, Keytip = "TW" } },
                { edtMarginLeft, new ElementUi { Label = R1.edtMarginLeft, Keytip = "ML" } },
                { edtMarginRight, new ElementUi { Label = R1.edtMarginRight, Keytip = "MR" } },
                { edtMarginTop, new ElementUi { Label = R1.edtMarginTop, Keytip = "MT" } },
                { edtMarginBottom, new ElementUi { Label = R1.edtMarginBottom, Keytip = "MB" } },
                { btnResetHorizontalMargin, new ElementUi { Label = R1.btnResetHorizontalMargin, Image = R2.TextboxResetMargin, Keytip = "MH" } },
                { btnResetVerticalMargin, new ElementUi { Label = R1.btnResetVerticalMargin, Image = R2.TextboxResetMargin, Keytip = "MV" } },
                // grpShapeSizeAndPosition
                { grpShapeSizeAndPosition, new ElementUi { Label = R1.grpShapeSizeAndPosition, Image = R2.SizeAndPosition } },
                { mnuShapeArrangement, new ElementUi { Label = R1.mnuShapeArrangement, Image = R2.ObjectArrangement_32, Keytip = "B" } },
                { btnLockShapeAspectRatio, new ElementUi { Label = R1.btnLockShapeAspectRatio, Image = R2.ObjectLockAspectRatio, Keytip = "L" } },
                { btnShapeScaleAnchor, new ElementUi { Label = R1.btnScaleAnchor_TopLeft, Image = R2.ScaleFromTopLeft, Keytip = "PA" } },
                { btnCopyShapeSize, new ElementUi { Label = R1.btnCopyShapeSize, Image = R2.Copy, Keytip = "SC" } },
                { btnPasteShapeSize, new ElementUi { Label = R1.btnPasteShapeSize, Image = R2.Paste, Keytip = "SP" } },
                { edtShapePositionX, new ElementUi { Label = R1.edtShapePositionX, Keytip = "PX" } },
                { edtShapePositionY, new ElementUi { Label = R1.edtShapePositionY, Keytip = "PY" } },
                { btnCopyShapePosition, new ElementUi { Label = R1.btnCopyShapePosition, Image = R2.Copy, Keytip = "PC" } },
                { btnPasteShapePosition, new ElementUi { Label = R1.btnPasteShapePosition, Image = R2.Paste, Keytip = "PP" } },
                // grpReplacePicture
                { grpReplacePicture, new ElementUi { Label = R1.grpReplacePicture, Image = R2.PictureChangeFromClipboard } },
                { btnReplaceWithClipboard, new ElementUi { Label = R1.btnReplaceWithClipboard, Image = R2.PictureChangeFromClipboard_32, Keytip = "TC" } },
                { btnReplaceWithFile, new ElementUi { Label = R1.btnReplaceWithFile, Image = R2.PictureChange, Keytip = "TF" } },
                { chkReserveOriginalSize, new ElementUi { Label = R1.chkReserveOriginalSize, Keytip = "TR" } },
                { chkReplaceToMiddle, new ElementUi { Label = R1.chkReplaceToMiddle, Keytip = "TM" } },
                // grpPictureSizeAndPosition
                { grpPictureSizeAndPosition, new ElementUi { Label = R1.grpPictureSizeAndPosition, Image = R2.SizeAndPosition } },
                { mnuPictureArrangement, new ElementUi { Label = R1.mnuPictureArrangement, Image = R2.ObjectArrangement_32, Keytip = "B" } },
                { btnResetPictureSize, new ElementUi { Label = R1.btnResetPictureSize, Image = R2.PictureResetSize_32, Keytip = "SR" } },
                { btnLockPictureAspectRatio, new ElementUi { Label = R1.btnLockPictureAspectRatio, Image = R2.ObjectLockAspectRatio, Keytip = "L" } },
                { btnPictureScaleAnchor, new ElementUi { Label = R1.btnScaleAnchor_TopLeft, Image = R2.ScaleFromTopLeft, Keytip = "PA" } },
                { btnCopyPictureSize, new ElementUi { Label = R1.btnCopyPictureSize, Image = R2.Copy, Keytip = "SC" } },
                { btnPastePictureSize, new ElementUi { Label = R1.btnPastePictureSize, Image = R2.Paste, Keytip = "SP" } },
                { edtPicturePositionX, new ElementUi { Label = R1.edtPicturePositionX, Keytip = "PX" } },
                { edtPicturePositionY, new ElementUi { Label = R1.edtPicturePositionY, Keytip = "PY" } },
                { btnCopyPicturePosition, new ElementUi { Label = R1.btnCopyPicturePosition, Image = R2.Copy, Keytip = "PC" } },
                { btnPastePicturePosition, new ElementUi { Label = R1.btnPastePicturePosition, Image = R2.Paste, Keytip = "PP" } },
                // grpVideoSizeAndPosition
                { grpVideoSizeAndPosition, new ElementUi { Label = R1.grpVideoSizeAndPosition, Image = R2.SizeAndPosition } },
                { mnuVideoArrangement, new ElementUi { Label = R1.mnuVideoArrangement, Image = R2.ObjectArrangement_32, Keytip = "B" } },
                { btnResetVideoSize, new ElementUi { Label = R1.btnResetVideoSize, Image = R2.PictureResetSize_32, Keytip = "SR" } },
                { btnLockVideoAspectRatio, new ElementUi { Label = R1.btnLockVideoAspectRatio, Image = R2.ObjectLockAspectRatio, Keytip = "SL" } }, // L
                { btnVideoScaleAnchor, new ElementUi { Label = R1.btnScaleAnchor_TopLeft, Image = R2.ScaleFromTopLeft, Keytip = "SF" } }, // PA
                { btnCopyVideoSize, new ElementUi { Label = R1.btnCopyVideoSize, Image = R2.Copy, Keytip = "SC" } },
                { btnPasteVideoSize, new ElementUi { Label = R1.btnPasteVideoSize, Image = R2.Paste, Keytip = "SP" } },
                { edtVideoPositionX, new ElementUi { Label = R1.edtVideoPositionX, Keytip = "SX" } }, // PX
                { edtVideoPositionY, new ElementUi { Label = R1.edtVideoPositionY, Keytip = "SY" } }, // PY
                { btnCopyVideoPosition, new ElementUi { Label = R1.btnCopyVideoPosition, Image = R2.Copy, Keytip = "SS" } }, // PC
                { btnPasteVideoPosition, new ElementUi { Label = R1.btnPasteVideoPosition, Image = R2.Paste, Keytip = "ST" } }, // PP
                // grpAudioSizeAndPosition
                { grpAudioSizeAndPosition, new ElementUi { Label = R1.grpAudioSizeAndPosition, Image = R2.SizeAndPosition } },
                { mnuAudioArrangement, new ElementUi { Label = R1.mnuAudioArrangement, Image = R2.ObjectArrangement_32, Keytip = "B" } },
                { btnResetAudioSize, new ElementUi { Label = R1.btnResetAudioSize, Image = R2.PictureResetSize_32, Keytip = "SR" } },
                { btnLockAudioAspectRatio, new ElementUi { Label = R1.btnLockAudioAspectRatio, Image = R2.ObjectLockAspectRatio, Keytip = "L" } },
                { btnAudioScaleAnchor, new ElementUi { Label = R1.btnScaleAnchor_TopLeft, Image = R2.ScaleFromTopLeft, Keytip = "PA" } },
                { btnCopyAudioSize, new ElementUi { Label = R1.btnCopyAudioSize, Image = R2.Copy, Keytip = "SC" } },
                { btnPasteAudioSize, new ElementUi { Label = R1.btnPasteAudioSize, Image = R2.Paste, Keytip = "SP" } },
                { edtAudioPositionX, new ElementUi { Label = R1.edtAudioPositionX, Keytip = "PX" } },
                { edtAudioPositionY, new ElementUi { Label = R1.edtAudioPositionY, Keytip = "PY" } },
                { btnCopyAudioPosition, new ElementUi { Label = R1.btnCopyAudioPosition, Image = R2.Copy, Keytip = "PC" } },
                { btnPasteAudioPosition, new ElementUi { Label = R1.btnPasteAudioPosition, Image = R2.Paste, Keytip = "PP" } },
                // grpTableSizeAndPosition
                { grpTableSizeAndPosition, new ElementUi { Label = R1.grpTableSizeAndPosition, Image = R2.SizeAndPosition } },
                { mnuTableArrangement, new ElementUi { Label = R1.mnuTableArrangement, Image = R2.ObjectArrangement_32, Keytip = "SB" } }, // B
                { btnLockTableAspectRatio, new ElementUi { Label = R1.btnLockTableAspectRatio, Image = R2.ObjectLockAspectRatio, Keytip = "SL" } }, // L
                { btnTableScaleAnchor, new ElementUi { Label = R1.btnScaleAnchor_TopLeft, Image = R2.ScaleFromTopLeft, Keytip = "SF" } }, // PA
                { btnCopyTableSize, new ElementUi { Label = R1.btnCopyTableSize, Image = R2.Copy, Keytip = "SC" } },
                { btnPasteTableSize, new ElementUi { Label = R1.btnPasteTableSize, Image = R2.Paste, Keytip = "SP" } },
                { edtTablePositionX, new ElementUi { Label = R1.edtTablePositionX, Keytip = "SX" } }, // PX
                { edtTablePositionY, new ElementUi { Label = R1.edtTablePositionY, Keytip = "SY" } }, // PY
                { btnCopyTablePosition, new ElementUi { Label = R1.btnCopyTablePosition, Image = R2.Copy, Keytip = "SS" } }, // PC
                { btnPasteTablePosition, new ElementUi { Label = R1.btnPasteTablePosition, Image = R2.Paste, Keytip = "ST" } }, // PP
                // grpChartSizeAndPosition
                { grpChartSizeAndPosition, new ElementUi { Label = R1.grpChartSizeAndPosition, Image = R2.SizeAndPosition } },
                { mnuChartArrangement, new ElementUi { Label = R1.mnuChartArrangement, Image = R2.ObjectArrangement_32, Keytip = "B" } },
                { btnLockChartAspectRatio, new ElementUi { Label = R1.btnLockChartAspectRatio, Image = R2.ObjectLockAspectRatio, Keytip = "L" } },
                { btnChartScaleAnchor, new ElementUi { Label = R1.btnScaleAnchor_TopLeft, Image = R2.ScaleFromTopLeft, Keytip = "PA" } },
                { btnCopyChartSize, new ElementUi { Label = R1.btnCopyChartSize, Image = R2.Copy, Keytip = "SC" } },
                { btnPasteChartSize, new ElementUi { Label = R1.btnPasteChartSize, Image = R2.Paste, Keytip = "SP" } },
                { edtChartPositionX, new ElementUi { Label = R1.edtChartPositionX, Keytip = "PX" } },
                { edtChartPositionY, new ElementUi { Label = R1.edtChartPositionY, Keytip = "PY" } },
                { btnCopyChartPosition, new ElementUi { Label = R1.btnCopyChartPosition, Image = R2.Copy, Keytip = "PC" } },
                { btnPasteChartPosition, new ElementUi { Label = R1.btnPasteChartPosition, Image = R2.Paste, Keytip = "PP" } },
                // grpSmartartSizeAndPosition
                { grpSmartartSizeAndPosition, new ElementUi { Label = R1.grpSmartartSizeAndPosition, Image = R2.SizeAndPosition } },
                { mnuSmartartArrangement, new ElementUi { Label = R1.mnuSmartartArrangement, Image = R2.ObjectArrangement_32, Keytip = "B" } },
                { btnLockSmartartAspectRatio, new ElementUi { Label = R1.btnLockSmartartAspectRatio, Image = R2.ObjectLockAspectRatio, Keytip = "L" } },
                { btnSmartartScaleAnchor, new ElementUi { Label = R1.btnScaleAnchor_TopLeft, Image = R2.ScaleFromTopLeft, Keytip = "PA" } },
                { btnCopySmartartSize, new ElementUi { Label = R1.btnCopySmartartSize, Image = R2.Copy, Keytip = "SC" } },
                { btnPasteSmartartSize, new ElementUi { Label = R1.btnPasteSmartartSize, Image = R2.Paste, Keytip = "SP" } },
                { edtSmartartPositionX, new ElementUi { Label = R1.edtSmartartPositionX, Keytip = "PX" } },
                { edtSmartartPositionY, new ElementUi { Label = R1.edtSmartartPositionY, Keytip = "PY" } },
                { btnCopySmartartPosition, new ElementUi { Label = R1.btnCopySmartartPosition, Image = R2.Copy, Keytip = "PC" } },
                { btnPasteSmartartPosition, new ElementUi { Label = R1.btnPasteSmartartPosition, Image = R2.Paste, Keytip = "PP" } },
                // mnuArrangement
                { mnuArrangement_sepAlignmentAndResizing, new ElementUi { Label = R1.mnuArrangement_sepAlignmentAndResizing } },
                { mnuArrangement_mnuAlignment, new ElementUi { Label = R1.mnuArrangement_mnuAlignment, Image = R2.ObjectArrangement } },
                { mnuArrangement_mnuResizing, new ElementUi { Label = R1.mnuArrangement_mnuResizing, Image = R2.ScaleSameWidth } },
                { mnuArrangement_mnuSnapping, new ElementUi { Label = R1.mnuArrangement_mnuSnapping, Image = R2.SnapLeftToRight } },
                { mnuArrangement_mnuRotation, new ElementUi { Label = R1.mnuArrangement_mnuRotation, Image = R2.ObjectRotateRight90 } },
                { mnuArrangement_sepLayerOrderAndGrouping, new ElementUi { Label = R1.mnuArrangement_sepLayerOrderAndGrouping } },
                { mnuArrangement_mnuLayerOrder, new ElementUi { Label = R1.mnuArrangement_mnuLayerOrder, Image = R2.ObjectSendToBack } },
                { mnuArrangement_mnuGrouping, new ElementUi { Label = R1.mnuArrangement_mnuGrouping, Image = R2.ObjectsGroup } },
                { mnuArrangement_sepObjectsInSlide, new ElementUi { Label = R1.mnuArrangement_sepObjectsInSlide } },
                { mnuArrangement_sepAddInSetting, new ElementUi { Label = R1.mnuArrangement_sepAddInSetting } }
            };
        }

        private readonly Dictionary<string, Dictionary<string, string>> _msoKeytips = new() {
            {
                grpWordArt, new() {
                    { "TextStylesGallery", "AQ" },
                    { "TextFillColorPicker", "AF" },
                    { "TextOutlineColorPicker", "AU" },
                    { "TextEffectsMenu", "AE" },
                    { "WordArtFormatDialog", "AG" }
                }
            },
            { grpArrange, new() { { "GridSettings", "DG" }, { "ObjectSizeAndPositionDialog", "HS" }, { "SelectionPane", "HP" } } },
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
