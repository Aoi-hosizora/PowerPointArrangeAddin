using System.Collections.Generic;
using System.Linq;
using Forms = System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointArrangeAddin.Helper;
using PowerPointArrangeAddin.Misc;

#nullable enable

namespace PowerPointArrangeAddin.Ribbon {

    public partial class ArrangeRibbon {

        private const string ArrangeRibbonXmlName = "PowerPointArrangeAddin.Ribbon.ArrangeRibbon.UI.xml";
        private const string ArrangeRibbonMenuXmlName = "PowerPointArrangeAddin.Ribbon.ArrangeRibbon.Menu.xml";

        private Office.IRibbonUI? _ribbon;

        public void Ribbon_Load(Office.IRibbonUI ribbonUi) {
            _ribbon = ribbonUi;
        }

        private delegate bool AvailabilityRule(PowerPoint.ShapeRange? shapeRange, int shapesCount, bool hasTextFrame);
        private readonly Dictionary<string, AvailabilityRule> _availabilityRules;

        private Dictionary<string, AvailabilityRule> GenerateNewAvailabilityRules() {
            return new Dictionary<string, AvailabilityRule> {
                // grpArrange
                { btnAlignLeft, (_, cnt, _) => cnt >= 1 },
                { btnAlignCenter, (_, cnt, _) => cnt >= 1 },
                { btnAlignRight, (_, cnt, _) => cnt >= 1 },
                { btnAlignTop, (_, cnt, _) => cnt >= 1 },
                { btnAlignMiddle, (_, cnt, _) => cnt >= 1 },
                { btnAlignBottom, (_, cnt, _) => cnt >= 1 },
                { btnDistributeHorizontal, (shapeRange, cnt, _) => cnt >= 1 && ArrangementHelper.IsDistributable(shapeRange, _alignRelativeFlag) },
                { btnDistributeVertical, (shapeRange, cnt, _) => cnt >= 1 && ArrangementHelper.IsDistributable(shapeRange, _alignRelativeFlag) },
                { btnAlignRelative, (_, cnt, _) => cnt >= 1 },
                { btnScaleSameWidth, (_, cnt, _) => cnt >= 2 },
                { btnScaleSameHeight, (_, cnt, _) => cnt >= 2 },
                { btnScaleSameSize, (_, cnt, _) => cnt >= 2 },
                { btnScaleAnchor, (_, _, _) => true },
                { btnExtendSameLeft, (_, cnt, _) => cnt >= 2 },
                { btnExtendSameRight, (_, cnt, _) => cnt >= 2 },
                { btnExtendSameTop, (_, cnt, _) => cnt >= 2 },
                { btnExtendSameBottom, (_, cnt, _) => cnt >= 2 },
                { btnSnapLeft, (_, cnt, _) => cnt >= 2 },
                { btnSnapRight, (_, cnt, _) => cnt >= 2 },
                { btnSnapTop, (_, cnt, _) => cnt >= 2 },
                { btnSnapBottom, (_, cnt, _) => cnt >= 2 },
                { btnMoveForward, (_, cnt, _) => cnt >= 1 },
                { btnMoveFront, (_, cnt, _) => cnt >= 1 },
                { btnMoveBackward, (_, cnt, _) => cnt >= 1 },
                { btnMoveBack, (_, cnt, _) => cnt >= 1 },
                { btnRotateRight90, (_, cnt, _) => cnt >= 1 },
                { btnRotateLeft90, (_, cnt, _) => cnt >= 1 },
                { btnFlipVertical, (_, cnt, _) => cnt >= 1 },
                { btnFlipHorizontal, (_, cnt, _) => cnt >= 1 },
                { btnGroup, (_, cnt, _) => cnt >= 2 },
                { btnUngroup, (shapeRange, cnt, _) => cnt >= 1 && ArrangementHelper.IsUngroupable(shapeRange) },
                { mnuArrangement, (_, _, _) => true },
                { btnAddInSetting, (_, _, _) => true },
                // mnuArrange
                { mnuArrangement_btnAlignRelative_ToObjects, (_, cnt, _) => cnt >= 2 },
                { mnuArrangement_btnAlignRelative_ToFirstObject, (_, cnt, _) => cnt >= 2 },
                { mnuArrangement_btnAlignRelative_ToSlide, (_, cnt, _) => cnt >= 1 },
                { mnuArrangement_btnScaleAnchor_FromTopLeft, (_, _, _) => true },
                { mnuArrangement_btnScaleAnchor_FromMiddle, (_, _, _) => true },
                { mnuArrangement_btnScaleAnchor_FromBottomRight, (_, _, _) => true },
                // grpTextbox
                { btnAutofitOff, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnAutoShrinkText, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnAutoResizeShape, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnWrapText, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginLeft, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginRight, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginTop, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginBottom, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnResetHorizontalMargin, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnResetVerticalMargin, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                // grpShapeSizeAndPosition
                { mnuShapeArrangement, (_, _, _) => true },
                { btnLockShapeAspectRatio, (_, cnt, _) => cnt >= 1 },
                { btnShapeScaleAnchor, (_, _, _) => true },
                { btnCopyShapeSize, (_, cnt, _) => cnt == 1 },
                { btnPasteShapeSize, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedSizeValue() },
                { edtShapePositionX, (_, cnt, _) => cnt >= 1 },
                { edtShapePositionY, (_, cnt, _) => cnt >= 1 },
                { btnCopyShapePosition, (_, cnt, _) => cnt == 1 },
                { btnPasteShapePosition, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedPositionValue() },
                // grpReplacePicture
                { btnReplaceWithClipboard, (_, cnt, _) => cnt >= 1 },
                { btnReplaceWithFile, (_, cnt, _) => cnt >= 1 },
                { chkReserveOriginalSize, (_, _, _) => true },
                { chkReplaceToMiddle, (_, _, _) => true },
                // grpPictureSizeAndPosition
                { mnuPictureArrangement, (_, _, _) => true },
                { btnResetPictureSize, (_, cnt, _) => cnt >= 1 },
                { btnLockPictureAspectRatio, (_, cnt, _) => cnt >= 1 },
                { btnPictureScaleAnchor, (_, _, _) => true },
                { btnCopyPictureSize, (_, cnt, _) => cnt == 1 },
                { btnPastePictureSize, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedSizeValue() },
                { edtPicturePositionX, (_, cnt, _) => cnt >= 1 },
                { edtPicturePositionY, (_, cnt, _) => cnt >= 1 },
                { btnCopyPicturePosition, (_, cnt, _) => cnt == 1 },
                { btnPastePicturePosition, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedPositionValue() },
                // grpVideoSizeAndPosition
                { mnuVideoArrangement, (_, _, _) => true },
                { btnResetVideoSize, (_, cnt, _) => cnt >= 1 },
                { btnLockVideoAspectRatio, (_, cnt, _) => cnt >= 1 },
                { btnVideoScaleAnchor, (_, _, _) => true },
                { btnCopyVideoSize, (_, cnt, _) => cnt == 1 },
                { btnPasteVideoSize, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedSizeValue() },
                { edtVideoPositionX, (_, cnt, _) => cnt >= 1 },
                { edtVideoPositionY, (_, cnt, _) => cnt >= 1 },
                { btnCopyVideoPosition, (_, cnt, _) => cnt == 1 },
                { btnPasteVideoPosition, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedPositionValue() },
                // grpAudioSizeAndPosition
                { mnuAudioArrangement, (_, _, _) => true },
                { btnResetAudioSize, (_, cnt, _) => cnt >= 1 },
                { btnLockAudioAspectRatio, (_, cnt, _) => cnt >= 1 },
                { btnAudioScaleAnchor, (_, _, _) => true },
                { btnCopyAudioSize, (_, cnt, _) => cnt == 1 },
                { btnPasteAudioSize, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedSizeValue() },
                { edtAudioPositionX, (_, cnt, _) => cnt >= 1 },
                { edtAudioPositionY, (_, cnt, _) => cnt >= 1 },
                { btnCopyAudioPosition, (_, cnt, _) => cnt == 1 },
                { btnPasteAudioPosition, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedPositionValue() },
                // grpTableSizeAndPosition
                { mnuTableArrangement, (_, _, _) => true },
                { btnLockTableAspectRatio, (_, cnt, _) => cnt >= 1 },
                { btnTableScaleAnchor, (_, _, _) => true },
                { btnCopyTableSize, (_, cnt, _) => cnt == 1 },
                { btnPasteTableSize, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedSizeValue() },
                { edtTablePositionX, (_, cnt, _) => cnt >= 1 },
                { edtTablePositionY, (_, cnt, _) => cnt >= 1 },
                { btnCopyTablePosition, (_, cnt, _) => cnt == 1 },
                { btnPasteTablePosition, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedPositionValue() },
                // grpChartSizeAndPosition
                { mnuChartArrangement, (_, _, _) => true },
                { btnLockChartAspectRatio, (_, cnt, _) => cnt >= 1 },
                { btnChartScaleAnchor, (_, _, _) => true },
                { btnCopyChartSize, (_, cnt, _) => cnt == 1 },
                { btnPasteChartSize, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedSizeValue() },
                { edtChartPositionX, (_, cnt, _) => cnt >= 1 },
                { edtChartPositionY, (_, cnt, _) => cnt >= 1 },
                { btnCopyChartPosition, (_, cnt, _) => cnt == 1 },
                { btnPasteChartPosition, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedPositionValue() },
                // grpSmartartSizeAndPosition
                { mnuSmartartArrangement, (_, _, _) => true },
                { btnLockSmartartAspectRatio, (_, cnt, _) => cnt >= 1 },
                { btnSmartartScaleAnchor, (_, _, _) => true },
                { btnCopySmartartSize, (_, cnt, _) => cnt == 1 },
                { btnPasteSmartartSize, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedSizeValue() },
                { edtSmartartPositionX, (_, cnt, _) => cnt >= 1 },
                { edtSmartartPositionY, (_, cnt, _) => cnt >= 1 },
                { btnCopySmartartPosition, (_, cnt, _) => cnt == 1 },
                { btnPasteSmartartPosition, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedPositionValue() }
            };
        }

        public ArrangeRibbon() {
            _ribbonElementUis = GenerateNewElementUis();
            _availabilityRules = GenerateNewAvailabilityRules();
        }

        public bool GetEnabled(Office.IRibbonControl ribbonControl) {
            var selection = SelectionGetter.GetSelection(onlyShapeRange: false);
            var shapesCount = selection.ShapeRange?.Count ?? 0;
            var hasTextFrame = selection.TextFrame != null;
            _availabilityRules.TryGetValue(ribbonControl.Id, out var checker);
            return checker?.Invoke(selection.ShapeRange, shapesCount, hasTextFrame) ?? true;
        }

        public bool GetControlVisible(Office.IRibbonControl ribbonControl) {
            var controls = new[] { grpArrange_separator2, bgpMoveLayers, bgpRotate, bgpGroupObjects, grpArrange_separator3, mnuArrangement, };
            if (controls.Contains(ribbonControl.Id)) {
                return !AddInSetting.Instance.LessButtonsForArrangementGroup;
            }
            return true;
        }

        public bool GetGroupVisible(Office.IRibbonControl ribbonControl) {
            return ribbonControl.Id switch {
                grpWordArt => AddInSetting.Instance.ShowWordArtGroup,
                grpArrange => true,
                grpTextbox => AddInSetting.Instance.ShowShapeTextboxGroup,
                grpShapeSizeAndPosition => AddInSetting.Instance.ShowShapeSizeAndPositionGroup,
                grpReplacePicture => AddInSetting.Instance.ShowReplacePictureGroup,
                grpPictureSizeAndPosition => AddInSetting.Instance.ShowPictureSizeAndPositionGroup,
                grpVideoSizeAndPosition => AddInSetting.Instance.ShowVideoSizeAndPositionGroup,
                grpAudioSizeAndPosition => AddInSetting.Instance.ShowAudioSizeAndPositionGroup,
                grpTableSizeAndPosition => AddInSetting.Instance.ShowTableSizeAndPositionGroup,
                grpChartSizeAndPosition => AddInSetting.Instance.ShowChartSizeAndPositionGroup,
                grpSmartartSizeAndPosition => AddInSetting.Instance.ShowSmartartSizeAndPositionGroup,
                _ => true
            };
        }

        public void InvalidateRibbon(bool onlyForDrag = false) {
            if (!onlyForDrag) {
                _ribbon?.Invalidate();
            } else {
                // currently callback that only for dragging to change the position is unavailable
                _ribbon?.InvalidateControl(edtShapePositionX);
                _ribbon?.InvalidateControl(edtShapePositionY);
            }
        }

        private PowerPoint.ShapeRange? GetShapeRange(int mustMoreThanOrEqualTo = 1) {
            var selection = SelectionGetter.GetSelection(onlyShapeRange: true);
            var shapeRange = selection.ShapeRange;
            if (shapeRange == null || shapeRange.Count < mustMoreThanOrEqualTo) {
                return null;
            }
            return shapeRange;
        }

        private (PowerPoint.TextFrame?, PowerPoint.TextFrame2?) GetTextFrame() {
            var selection = SelectionGetter.GetSelection(onlyShapeRange: false);
            return (selection.TextFrame, selection.TextFrame2);
        }

        public void BtnAlign_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            Office.MsoAlignCmd? cmd = ribbonControl.Id switch {
                btnAlignLeft => Office.MsoAlignCmd.msoAlignLefts,
                btnAlignCenter => Office.MsoAlignCmd.msoAlignCenters,
                btnAlignRight => Office.MsoAlignCmd.msoAlignRights,
                btnAlignTop => Office.MsoAlignCmd.msoAlignTops,
                btnAlignMiddle => Office.MsoAlignCmd.msoAlignMiddles,
                btnAlignBottom => Office.MsoAlignCmd.msoAlignBottoms,
                _ => null
            };
            ArrangementHelper.Align(shapeRange, cmd, _alignRelativeFlag);
        }

        public void BtnDistribute_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            Office.MsoDistributeCmd? cmd = ribbonControl.Id switch {
                btnDistributeHorizontal => Office.MsoDistributeCmd.msoDistributeHorizontally,
                btnDistributeVertical => Office.MsoDistributeCmd.msoDistributeVertically,
                _ => null
            };
            ArrangementHelper.Distribute(shapeRange, cmd, _alignRelativeFlag);
        }

        // This flag is used to adjust alignment relative behavior, is used by BtnAlign_Click and BtnDistribute_Click.
        private ArrangementHelper.AlignRelativeFlag _alignRelativeFlag = ArrangementHelper.AlignRelativeFlag.RelativeToObjects;

        public void BtnAlignRelative_Click(Office.IRibbonControl ribbonControl, bool _ = false) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null || shapeRange.Count <= 1) {
                return; // not change if no more than 1 shape is selected
            }
            if (!ribbonControl.Id.Contains("_To")) {
                _alignRelativeFlag = _alignRelativeFlag switch {
                    ArrangementHelper.AlignRelativeFlag.RelativeToObjects => ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject,
                    ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject => ArrangementHelper.AlignRelativeFlag.RelativeToSlide,
                    ArrangementHelper.AlignRelativeFlag.RelativeToSlide => ArrangementHelper.AlignRelativeFlag.RelativeToObjects,
                    _ => ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject
                };
            } else {
                _alignRelativeFlag = ribbonControl.Id switch {
                    mnuArrangement_btnAlignRelative_ToObjects => ArrangementHelper.AlignRelativeFlag.RelativeToObjects,
                    mnuArrangement_btnAlignRelative_ToFirstObject => ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject,
                    mnuArrangement_btnAlignRelative_ToSlide => ArrangementHelper.AlignRelativeFlag.RelativeToSlide,
                    _ => ArrangementHelper.AlignRelativeFlag.RelativeToObjects,
                };
            }

            var mso = _alignRelativeFlag == ArrangementHelper.AlignRelativeFlag.RelativeToSlide
                ? "ObjectsAlignRelativeToContainerSmart"
                : "ObjectsAlignSelectedSmart";
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso(mso); // consist with mso

            _ribbon?.InvalidateControl(btnAlignRelative);
            _ribbon?.InvalidateControl(btnDistributeHorizontal);
            _ribbon?.InvalidateControl(btnDistributeVertical);
            _ribbon?.InvalidateControl(mnuArrangement_btnAlignRelative_ToObjects);
            _ribbon?.InvalidateControl(mnuArrangement_btnAlignRelative_ToFirstObject);
            _ribbon?.InvalidateControl(mnuArrangement_btnAlignRelative_ToSlide);
        }

        public bool BtnAlignRelative_GetPressed(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null || shapeRange.Count <= 1) {
                return ribbonControl.Id == mnuArrangement_btnAlignRelative_ToSlide; // if no more than 1 shape is selected 
            }
            return ribbonControl.Id switch {
                mnuArrangement_btnAlignRelative_ToObjects => _alignRelativeFlag == ArrangementHelper.AlignRelativeFlag.RelativeToObjects,
                mnuArrangement_btnAlignRelative_ToFirstObject => _alignRelativeFlag == ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject,
                mnuArrangement_btnAlignRelative_ToSlide => _alignRelativeFlag == ArrangementHelper.AlignRelativeFlag.RelativeToSlide,
                _ => false
            };
        }

        public string BtnAlignRelative_GetLabel(Office.IRibbonControl ribbonControl) {
            if (!ribbonControl.Id.Contains("_To")) {
                var shapeRange = GetShapeRange();
                if (shapeRange?.Count == 1) {
                    return ArrangeRibbonResources.btnAlignRelative_ToSlide; // when single shape is selected
                }
                return _alignRelativeFlag switch {
                    ArrangementHelper.AlignRelativeFlag.RelativeToObjects => ArrangeRibbonResources.btnAlignRelative_ToObjects,
                    ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject => ArrangeRibbonResources.btnAlignRelative_ToFirstObject,
                    ArrangementHelper.AlignRelativeFlag.RelativeToSlide => ArrangeRibbonResources.btnAlignRelative_ToSlide,
                    _ => ArrangeRibbonResources.btnAlignRelative_ToObjects
                };
            }
            return ribbonControl.Id switch {
                mnuArrangement_btnAlignRelative_ToObjects => ArrangeRibbonResources.btnAlignRelative_ToObjects,
                mnuArrangement_btnAlignRelative_ToFirstObject => ArrangeRibbonResources.btnAlignRelative_ToFirstObject,
                mnuArrangement_btnAlignRelative_ToSlide => ArrangeRibbonResources.btnAlignRelative_ToSlide,
                _ => ArrangeRibbonResources.btnAlignRelative_ToObjects
            };
        }

        public System.Drawing.Image BtnAlignRelative_GetImage(Office.IRibbonControl ribbonControl) {
            if (!ribbonControl.Id.Contains("_To")) {
                var shapeRange = GetShapeRange();
                if (shapeRange?.Count == 1) {
                    return Properties.Resources.AlignRelativeToSlide; // when single shape is selected
                }
                return _alignRelativeFlag switch {
                    ArrangementHelper.AlignRelativeFlag.RelativeToObjects => Properties.Resources.AlignRelativeToObjects,
                    ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject => Properties.Resources.AlignRelativeToFirstObject,
                    ArrangementHelper.AlignRelativeFlag.RelativeToSlide => Properties.Resources.AlignRelativeToSlide,
                    _ => Properties.Resources.AlignRelativeToObjects
                };
            }
            return ribbonControl.Id switch {
                mnuArrangement_btnAlignRelative_ToObjects => Properties.Resources.AlignRelativeToObjects,
                mnuArrangement_btnAlignRelative_ToFirstObject => Properties.Resources.AlignRelativeToFirstObject,
                mnuArrangement_btnAlignRelative_ToSlide => Properties.Resources.AlignRelativeToSlide,
                _ => Properties.Resources.AlignRelativeToObjects
            };
        }

        public void BtnScale_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.ScaleSizeCmd? cmd = ribbonControl.Id switch {
                btnScaleSameWidth => ArrangementHelper.ScaleSizeCmd.SameWidth,
                btnScaleSameHeight => ArrangementHelper.ScaleSizeCmd.SameHeight,
                btnScaleSameSize => ArrangementHelper.ScaleSizeCmd.SameSize,
                _ => null
            };
            ArrangementHelper.ScaleSize(shapeRange, cmd, _scaleFromFlag);
        }

        // This flag is used to control scale and size behavior, is used by BtnScale_Click, BtnCopyAndPasteSize_Click and BtnResetMediaSize_Click.
        private Office.MsoScaleFrom _scaleFromFlag = Office.MsoScaleFrom.msoScaleFromTopLeft;

        public void BtnScaleAnchor_Click(Office.IRibbonControl ribbonControl, bool _ = false) {
            if (!ribbonControl.Id.Contains("_From")) {
                _scaleFromFlag = _scaleFromFlag switch {
                    Office.MsoScaleFrom.msoScaleFromTopLeft => Office.MsoScaleFrom.msoScaleFromMiddle,
                    Office.MsoScaleFrom.msoScaleFromMiddle => Office.MsoScaleFrom.msoScaleFromBottomRight,
                    Office.MsoScaleFrom.msoScaleFromBottomRight => Office.MsoScaleFrom.msoScaleFromTopLeft,
                    _ => Office.MsoScaleFrom.msoScaleFromTopLeft
                };
            } else {
                _scaleFromFlag = ribbonControl.Id switch {
                    mnuArrangement_btnScaleAnchor_FromTopLeft => Office.MsoScaleFrom.msoScaleFromTopLeft,
                    mnuArrangement_btnScaleAnchor_FromMiddle => Office.MsoScaleFrom.msoScaleFromMiddle,
                    mnuArrangement_btnScaleAnchor_FromBottomRight => Office.MsoScaleFrom.msoScaleFromBottomRight,
                    _ => _scaleFromFlag
                };
            }

            _ribbon?.InvalidateControl(btnScaleAnchor);
            _ribbon?.InvalidateControl(mnuArrangement_btnScaleAnchor_FromTopLeft);
            _ribbon?.InvalidateControl(mnuArrangement_btnScaleAnchor_FromMiddle);
            _ribbon?.InvalidateControl(mnuArrangement_btnScaleAnchor_FromBottomRight);
            _ribbon?.InvalidateControl(btnShapeScaleAnchor);
            _ribbon?.InvalidateControl(btnPictureScaleAnchor);
            _ribbon?.InvalidateControl(btnVideoScaleAnchor);
            _ribbon?.InvalidateControl(btnAudioScaleAnchor);
            _ribbon?.InvalidateControl(btnTableScaleAnchor);
            _ribbon?.InvalidateControl(btnChartScaleAnchor);
            _ribbon?.InvalidateControl(btnSmartartScaleAnchor);
        }

        public bool BtnScaleAnchor_GetPressed(Office.IRibbonControl ribbonControl) {
            return ribbonControl.Id switch {
                mnuArrangement_btnScaleAnchor_FromTopLeft => _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromTopLeft,
                mnuArrangement_btnScaleAnchor_FromMiddle => _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle,
                mnuArrangement_btnScaleAnchor_FromBottomRight => _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromBottomRight,
                _ => false
            };
        }

        public string BtnScaleAnchor_GetLabel(Office.IRibbonControl ribbonControl) {
            if (!ribbonControl.Id.Contains("_From")) {
                return _scaleFromFlag switch {
                    Office.MsoScaleFrom.msoScaleFromTopLeft => ArrangeRibbonResources.btnScaleAnchor_TopLeft,
                    Office.MsoScaleFrom.msoScaleFromMiddle => ArrangeRibbonResources.btnScaleAnchor_Middle,
                    Office.MsoScaleFrom.msoScaleFromBottomRight => ArrangeRibbonResources.btnScaleAnchor_BottomRight,
                    _ => ArrangeRibbonResources.btnScaleAnchor_TopLeft
                };
            }
            return ribbonControl.Id switch {
                mnuArrangement_btnScaleAnchor_FromTopLeft => ArrangeRibbonResources.btnScaleAnchor_TopLeft,
                mnuArrangement_btnScaleAnchor_FromMiddle => ArrangeRibbonResources.btnScaleAnchor_Middle,
                mnuArrangement_btnScaleAnchor_FromBottomRight => ArrangeRibbonResources.btnScaleAnchor_BottomRight,
                _ => ArrangeRibbonResources.btnScaleAnchor_TopLeft
            };
        }

        public System.Drawing.Image BtnScaleAnchor_GetImage(Office.IRibbonControl ribbonControl) {
            if (!ribbonControl.Id.Contains("_From")) {
                return _scaleFromFlag switch {
                    Office.MsoScaleFrom.msoScaleFromTopLeft => Properties.Resources.ScaleFromTopLeft,
                    Office.MsoScaleFrom.msoScaleFromMiddle => Properties.Resources.ScaleFromMiddle,
                    Office.MsoScaleFrom.msoScaleFromBottomRight => Properties.Resources.ScaleFromBottomRight,
                    _ => Properties.Resources.ScaleFromTopLeft
                };
            }
            return ribbonControl.Id switch {
                mnuArrangement_btnScaleAnchor_FromTopLeft => Properties.Resources.ScaleFromTopLeft,
                mnuArrangement_btnScaleAnchor_FromMiddle => Properties.Resources.ScaleFromMiddle,
                mnuArrangement_btnScaleAnchor_FromBottomRight => Properties.Resources.ScaleFromBottomRight,
                _ => Properties.Resources.ScaleFromTopLeft
            };
        }

        public void BtnExtend_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.ExtendSizeCmd? cmd = ribbonControl.Id switch {
                btnExtendSameLeft => ArrangementHelper.ExtendSizeCmd.ExtendToLeft,
                btnExtendSameRight => ArrangementHelper.ExtendSizeCmd.ExtendToRight,
                btnExtendSameTop => ArrangementHelper.ExtendSizeCmd.ExtendToTop,
                btnExtendSameBottom => ArrangementHelper.ExtendSizeCmd.ExtendToBottom,
                _ => null
            };
            ArrangementHelper.ExtendSize(shapeRange, cmd);
        }

        public void BtnSnap_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.SnapCmd? cmd = ribbonControl.Id switch {
                btnSnapLeft => ArrangementHelper.SnapCmd.SnapLeftToRight,
                btnSnapRight => ArrangementHelper.SnapCmd.SnapRightToLeft,
                btnSnapTop => ArrangementHelper.SnapCmd.SnapTopToBottom,
                btnSnapBottom => ArrangementHelper.SnapCmd.SnapBottomToTop,
                _ => null
            };
            ArrangementHelper.Snap(shapeRange, cmd);
        }

        public void BtnMove_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            Office.MsoZOrderCmd? cmd = ribbonControl.Id switch {
                btnMoveForward => Office.MsoZOrderCmd.msoBringForward,
                btnMoveBackward => Office.MsoZOrderCmd.msoSendBackward,
                btnMoveFront => Office.MsoZOrderCmd.msoBringToFront,
                btnMoveBack => Office.MsoZOrderCmd.msoSendToBack,
                _ => null
            };
            ArrangementHelper.LayerMove(shapeRange, cmd);
        }

        public void BtnRotate_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.RotateCmd? cmd = ribbonControl.Id switch {
                btnRotateLeft90 => ArrangementHelper.RotateCmd.RotateLeft90,
                btnRotateRight90 => ArrangementHelper.RotateCmd.RotateRight90,
                _ => null
            };
            ArrangementHelper.Rotate(shapeRange, cmd);
        }

        public void BtnFlip_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            Office.MsoFlipCmd? cmd = ribbonControl.Id switch {
                btnFlipHorizontal => Office.MsoFlipCmd.msoFlipHorizontal,
                btnFlipVertical => Office.MsoFlipCmd.msoFlipVertical,
                _ => null
            };
            ArrangementHelper.Flip(shapeRange, cmd);
        }

        public void BtnGroup_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.GroupCmd? cmd = ribbonControl.Id switch {
                btnGroup => ArrangementHelper.GroupCmd.Group,
                btnUngroup => ArrangementHelper.GroupCmd.Ungroup,
                _ => null
            };
            ArrangementHelper.Group(shapeRange, cmd, () => _ribbon?.Invalidate());
        }

        public void BtnAddInSetting_Click(Office.IRibbonControl _) {
            var oldLanguage = AddInSetting.Instance.Language;
            var dlg = new Dialog.SettingDialog();
            var result = dlg.ShowDialog();
            if (result != Forms.DialogResult.OK) {
                return;
            }
            if (AddInSetting.Instance.Language != oldLanguage) {
                // include updating elements and invalidating ribbon
                AddInLanguageChanger.ChangeLanguage(AddInSetting.Instance.Language);
            } else {
                _ribbon?.Invalidate(); // just invalidate ribbon
            }
        }

        public void BtnAutofit_Click(Office.IRibbonControl ribbonControl, bool _) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            TextboxHelper.TextboxStatusCmd? cmd = ribbonControl.Id switch {
                btnAutofitOff => TextboxHelper.TextboxStatusCmd.AutofitOff,
                btnAutoShrinkText => TextboxHelper.TextboxStatusCmd.AutoShrinkText,
                btnAutoResizeShape => TextboxHelper.TextboxStatusCmd.AutoResizeShape,
                _ => null
            };
            TextboxHelper.ChangeAutofitStatus(textFrame, cmd, () => {
                _ribbon?.InvalidateControl(btnAutofitOff);
                _ribbon?.InvalidateControl(btnAutoShrinkText);
                _ribbon?.InvalidateControl(btnAutoResizeShape);
            });
        }

        public bool BtnAutofit_GetPressed(Office.IRibbonControl ribbonControl) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return false;
            }
            TextboxHelper.TextboxStatusCmd? cmd = ribbonControl.Id switch {
                btnAutofitOff => TextboxHelper.TextboxStatusCmd.AutofitOff,
                btnAutoShrinkText => TextboxHelper.TextboxStatusCmd.AutoShrinkText,
                btnAutoResizeShape => TextboxHelper.TextboxStatusCmd.AutoResizeShape,
                _ => null
            };
            return TextboxHelper.GetAutofitStatus(textFrame, cmd);
        }

        public void BtnWrapText_Click(Office.IRibbonControl _, bool __) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            var cmd = TextboxHelper.TextboxStatusCmd.WrapTextOnOff;
            TextboxHelper.ChangeAutofitStatus(textFrame, cmd, () => {
                _ribbon?.InvalidateControl(btnWrapText);
            });
        }

        public bool BtnWrapText_GetPressed(Office.IRibbonControl _) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return false;
            }
            var cmd = TextboxHelper.TextboxStatusCmd.WrapTextOnOff;
            return TextboxHelper.GetAutofitStatus(textFrame, cmd);
        }

        public void BtnResetMargin_Click(Office.IRibbonControl ribbonControl) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            TextboxHelper.ResetMarginCmd? cmd = ribbonControl.Id switch {
                btnResetHorizontalMargin => TextboxHelper.ResetMarginCmd.Horizontal,
                btnResetVerticalMargin => TextboxHelper.ResetMarginCmd.Vertical,
                _ => null
            };
            TextboxHelper.ResetMargin(textFrame, cmd, () => {
                _ribbon?.InvalidateControl(edtMarginLeft);
                _ribbon?.InvalidateControl(edtMarginRight);
                _ribbon?.InvalidateControl(edtMarginTop);
                _ribbon?.InvalidateControl(edtMarginBottom);
            });
        }

        public void EdtMargin_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            TextboxHelper.MarginKind? kind = ribbonControl.Id switch {
                edtMarginLeft => TextboxHelper.MarginKind.Left,
                edtMarginRight => TextboxHelper.MarginKind.Right,
                edtMarginTop => TextboxHelper.MarginKind.Top,
                edtMarginBottom => TextboxHelper.MarginKind.Bottom,
                _ => null
            };
            TextboxHelper.ChangeMarginOfString(textFrame, kind, text, () => {
                _ribbon?.InvalidateControl(ribbonControl.Id);
            });
        }

        public string EdtMargin_GetText(Office.IRibbonControl ribbonControl) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return "";
            }
            TextboxHelper.MarginKind? kind = ribbonControl.Id switch {
                edtMarginLeft => TextboxHelper.MarginKind.Left,
                edtMarginRight => TextboxHelper.MarginKind.Right,
                edtMarginTop => TextboxHelper.MarginKind.Top,
                edtMarginBottom => TextboxHelper.MarginKind.Bottom,
                _ => null
            };
            return TextboxHelper.GetMarginOfString(textFrame, kind).Item1;
        }

        public void BtnLockAspectRatio_Click(Office.IRibbonControl ribbonControl, bool pressed) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            var cmd = SizeAndPositionHelper.LockAspectRatioCmd.Toggle;
            SizeAndPositionHelper.ToggleLockAspectRatio(shapeRange, cmd, () => {
                _ribbon?.InvalidateControl(ribbonControl.Id);
            });
        }

        public bool BtnLockAspectRatio_GetPressed(Office.IRibbonControl _) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return false;
            }
            return SizeAndPositionHelper.GetAspectRatioIsLocked(shapeRange);
        }

        public void BtnCopyAndPasteSize_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.CopyAndPasteCmd? cmd = ribbonControl.Id switch {
                btnCopyShapeSize or btnCopyPictureSize or btnCopyVideoSize or btnCopyAudioSize
                    or btnCopyTableSize or btnCopyChartSize or btnCopySmartartSize => SizeAndPositionHelper.CopyAndPasteCmd.Copy,
                btnPasteShapeSize or btnPastePictureSize or btnPasteVideoSize or btnPasteAudioSize
                    or btnPasteTableSize or btnPasteChartSize or btnPasteSmartartSize => SizeAndPositionHelper.CopyAndPasteCmd.Paste,
                _ => null
            };
            SizeAndPositionHelper.CopyAndPasteSize(shapeRange, cmd, _scaleFromFlag, () => {
                _ribbon?.InvalidateControl(btnPasteShapeSize);
                _ribbon?.InvalidateControl(btnPastePictureSize);
                _ribbon?.InvalidateControl(btnPasteVideoSize);
                _ribbon?.InvalidateControl(btnPasteAudioSize);
                _ribbon?.InvalidateControl(btnPasteTableSize);
                _ribbon?.InvalidateControl(btnPasteChartSize);
                _ribbon?.InvalidateControl(btnPasteSmartartSize);
            });
        }

        public void EdtPosition_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.PositionKind? kind = ribbonControl.Id switch {
                edtShapePositionX or edtPicturePositionX or edtVideoPositionX or edtAudioPositionX
                    or edtTablePositionX or edtChartPositionX or edtSmartartPositionX => SizeAndPositionHelper.PositionKind.X,
                edtShapePositionY or edtPicturePositionY or edtVideoPositionY or edtAudioPositionY
                    or edtTablePositionY or edtChartPositionY or edtSmartartPositionY => SizeAndPositionHelper.PositionKind.Y,
                _ => null
            };
            SizeAndPositionHelper.ChangePositionOfString(shapeRange, kind, text, () => {
                _ribbon?.InvalidateControl(ribbonControl.Id);
            });
        }

        public string EdtPosition_GetText(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return "";
            }
            SizeAndPositionHelper.PositionKind? kind = ribbonControl.Id switch {
                edtShapePositionX or edtPicturePositionX or edtVideoPositionX or edtAudioPositionX
                    or edtTablePositionX or edtChartPositionX or edtSmartartPositionX => SizeAndPositionHelper.PositionKind.X,
                edtShapePositionY or edtPicturePositionY or edtVideoPositionY or edtAudioPositionY
                    or edtTablePositionY or edtChartPositionY or edtSmartartPositionY => SizeAndPositionHelper.PositionKind.Y,
                _ => null
            };
            return SizeAndPositionHelper.GetPositionOfString(shapeRange, kind).Item1;
        }

        public void BtnCopyAndPastePosition_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.CopyAndPasteCmd? cmd = ribbonControl.Id switch {
                btnCopyShapePosition or btnCopyPicturePosition or btnCopyVideoPosition or btnCopyAudioPosition
                    or btnCopyTablePosition or btnCopyChartPosition or btnCopySmartartPosition => SizeAndPositionHelper.CopyAndPasteCmd.Copy,
                btnPasteShapePosition or btnPastePicturePosition or btnPasteVideoPosition or btnPasteAudioPosition
                    or btnPasteTablePosition or btnPasteChartPosition or btnPasteSmartartPosition => SizeAndPositionHelper.CopyAndPasteCmd.Paste,
                _ => null
            };
            SizeAndPositionHelper.CopyAndPastePosition(shapeRange, cmd, () => {
                var controlIds = new[] {
                    btnPasteShapePosition, edtShapePositionX, edtShapePositionY,
                    btnPastePicturePosition, edtPicturePositionX, edtPicturePositionY,
                    btnPasteVideoPosition, edtVideoPositionX, edtVideoPositionY,
                    btnPasteAudioPosition, edtAudioPositionX, edtAudioPositionY,
                    btnPasteTablePosition, edtTablePositionX, edtTablePositionY,
                    btnPasteChartPosition, edtChartPositionX, edtChartPositionY,
                    btnPasteSmartartPosition, edtSmartartPositionX, edtSmartartPositionY,
                };
                foreach (var controlId in controlIds) {
                    _ribbon?.InvalidateControl(controlId);
                }
            });
        }

        public void BtnResetMediaSize_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.ResetMediaSize(shapeRange, _scaleFromFlag, () => {
                _ribbon?.InvalidateControl(edtPicturePositionX);
                _ribbon?.InvalidateControl(edtPicturePositionY);
                _ribbon?.InvalidateControl(edtVideoPositionX);
                _ribbon?.InvalidateControl(edtVideoPositionY);
                _ribbon?.InvalidateControl(edtAudioPositionX);
                _ribbon?.InvalidateControl(edtAudioPositionY);
            });
        }

        // This flag is used for picture replacing related callbacks, that is BtnReplacePicture_Click.
        private ReplacePictureHelper.ReplacePictureFlag _replacePictureFlag =
            ReplacePictureHelper.ReplacePictureFlag.ReplaceToMiddle | ReplacePictureHelper.ReplacePictureFlag.ReserveOriginalSize;

        public void ChkReserveOriginalSize_Click(Office.IRibbonControl ribbonControl, bool _) {
            if (!ChkReserveOriginalSize_GetPressed(ribbonControl)) {
                _replacePictureFlag |= ReplacePictureHelper.ReplacePictureFlag.ReserveOriginalSize;
            } else {
                _replacePictureFlag &= ~ReplacePictureHelper.ReplacePictureFlag.ReserveOriginalSize;
            }
            _ribbon?.InvalidateControl(ribbonControl.Id);
        }

        public bool ChkReserveOriginalSize_GetPressed(Office.IRibbonControl _) {
            return (_replacePictureFlag & ReplacePictureHelper.ReplacePictureFlag.ReserveOriginalSize) != 0;
        }

        public void ChkReplaceToMiddle_Click(Office.IRibbonControl ribbonControl, bool _) {
            if (!ChkReplaceToMiddle_GetPressed(ribbonControl)) {
                _replacePictureFlag |= ReplacePictureHelper.ReplacePictureFlag.ReplaceToMiddle;
            } else {
                _replacePictureFlag &= ~ReplacePictureHelper.ReplacePictureFlag.ReplaceToMiddle;
            }
            _ribbon?.InvalidateControl(ribbonControl.Id);
        }

        public bool ChkReplaceToMiddle_GetPressed(Office.IRibbonControl _) {
            return (_replacePictureFlag & ReplacePictureHelper.ReplacePictureFlag.ReplaceToMiddle) != 0;
        }

        public void BtnReplacePicture_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ReplacePictureHelper.ReplacePictureCmd? cmd = ribbonControl.Id switch {
                btnReplaceWithClipboard => ReplacePictureHelper.ReplacePictureCmd.WithClipboard,
                btnReplaceWithFile => ReplacePictureHelper.ReplacePictureCmd.WithFile,
                _ => null
            };
            ReplacePictureHelper.ReplacePicture(shapeRange, cmd, _replacePictureFlag, () => _ribbon?.Invalidate());
        }

    }

}
