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

        public ArrangeRibbon() {
            (_ribbonElementUis, _ribbonElementUiSpecials) = GenerateNewElementUis();
            _availabilityRules = GenerateNewAvailabilityRules();
        }

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
                // grpTextbox
                { btnAutofitOff, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnAutoShrinkText, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnAutoResizeShape, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnWrapText, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnResetHorizontalMargin, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginLeft, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginRight, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnResetVerticalMargin, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginTop, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginBottom, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                // grpReplacePicture
                { btnReplaceWithClipboard, (_, cnt, _) => cnt >= 1 },
                { btnReplaceWithFile, (_, cnt, _) => cnt >= 1 },
                { chkReserveOriginalSize, (_, _, _) => true },
                { chkReplaceToMiddle, (_, _, _) => true },
                // grpSizeAndPosition
                { btnResetSize, (shapeRange, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsSizeResettable(shapeRange) },
                { btnLockAspectRatio, (_, cnt, _) => cnt >= 1 },
                { btnCopySize, (_, cnt, _) => cnt == 1 },
                { btnPasteSize, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedSizeValue() },
                { edtPositionX, (_, cnt, _) => cnt >= 1 },
                { edtPositionY, (_, cnt, _) => cnt >= 1 },
                { btnCopyPosition, (_, cnt, _) => cnt == 1 },
                { btnPastePosition, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedPositionValue() },
                // mnuArrange
                { btnAlignRelative_ToObjects, (_, cnt, _) => cnt >= 2 },
                { btnAlignRelative_ToFirstObject, (_, cnt, _) => cnt >= 2 },
                { btnAlignRelative_ToSlide, (_, cnt, _) => cnt >= 1 },
                { btnScaleAnchor_FromTopLeft, (_, _, _) => true },
                { btnScaleAnchor_FromMiddle, (_, _, _) => true },
                { btnScaleAnchor_FromBottomRight, (_, _, _) => true },
                // ====================================================================
                { btnSizeAndPosition, (_, cnt, _) => cnt >= 1 },
            };
        }

        public bool GetEnabled(Office.IRibbonControl ribbonControl) {
            var selection = SelectionGetter.GetSelection(onlyShapeRange: false);
            var shapesCount = selection.ShapeRange?.Count ?? 0;
            var hasTextFrame = selection.TextFrame != null;
            _availabilityRules.TryGetValue(ribbonControl.Id(), out var checker);
            return checker?.Invoke(selection.ShapeRange, shapesCount, hasTextFrame) ?? true;
        }

        public bool GetControlVisible(Office.IRibbonControl ribbonControl) {
            var arrangementControls = new[] { sepMoveLayers, bgpMoveLayers, bgpRotate, bgpGroupObjects, sepArrangement, mnuArrangement };
            if (arrangementControls.Contains(ribbonControl.Id()) && ribbonControl.Group() == grpArrange) {
                return !AddInSetting.Instance.LessButtonsForArrangementGroup;
            }
            var textboxControls = new[] { sepHorizontalMargin, bgpHorizontalMargin, edtMarginLeft, edtMarginRight, sepVerticalMargin, bgpVerticalMargin, edtMarginTop, edtMarginBottom };
            if (textboxControls.Contains(ribbonControl.Id()) && ribbonControl.Group() == grpTextbox) {
                return !AddInSetting.Instance.HideMarginSettingForTextboxGroup;
            }
            return true;
        }

        public bool GetGroupVisible(Office.IRibbonControl ribbonControl) {
            return ribbonControl.Id() switch {
                grpWordArt => AddInSetting.Instance.ShowWordArtGroup,
                grpArrange => true,
                grpTextbox => AddInSetting.Instance.ShowShapeTextboxGroup,
                grpReplacePicture => AddInSetting.Instance.ShowReplacePictureGroup,
                grpShapeSizeAndPosition => AddInSetting.Instance.ShowShapeSizeAndPositionGroup2,
                grpPictureSizeAndPosition => AddInSetting.Instance.ShowPictureSizeAndPositionGroup2,
                grpVideoSizeAndPosition => AddInSetting.Instance.ShowVideoSizeAndPositionGroup2,
                grpAudioSizeAndPosition => AddInSetting.Instance.ShowAudioSizeAndPositionGroup2,
                grpTableSizeAndPosition => AddInSetting.Instance.ShowTableSizeAndPositionGroup2,
                grpChartSizeAndPosition => AddInSetting.Instance.ShowChartSizeAndPositionGroup2,
                grpSmartartSizeAndPosition => AddInSetting.Instance.ShowSmartartSizeAndPositionGroup2,
                _ => true
            };
        }

        public void InvalidateRibbon(bool onlyForDrag = false) {
            if (!onlyForDrag) {
                _ribbon?.Invalidate();
            } else {
                // currently callback that only for dragging to change the position is unavailable
                _ribbon?.InvalidateControl(edtPositionX, _sizeAndPositionGroups);
                _ribbon?.InvalidateControl(edtPositionY, _sizeAndPositionGroups);
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
            Office.MsoAlignCmd? cmd = ribbonControl.Id() switch {
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
            Office.MsoDistributeCmd? cmd = ribbonControl.Id() switch {
                btnDistributeHorizontal => Office.MsoDistributeCmd.msoDistributeHorizontally,
                btnDistributeVertical => Office.MsoDistributeCmd.msoDistributeVertically,
                _ => null
            };
            ArrangementHelper.Distribute(shapeRange, cmd, _alignRelativeFlag);
        }

        // This flag is used to adjust alignment relative behavior, is used by BtnAlign_Click and BtnDistribute_Click.
        private ArrangementHelper.AlignRelativeFlag _alignRelativeFlag = ArrangementHelper.AlignRelativeFlag.RelativeToObjects;

        private bool IsOptionRibbonButton(Office.IRibbonControl ribbonControl) {
            return ribbonControl.Id().Contains("_To") || ribbonControl.Id().Contains("_From"); // just check by id
        }

        public void BtnAlignRelative_Click(Office.IRibbonControl ribbonControl, bool _ = false) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null || shapeRange.Count <= 1) {
                _ribbon?.InvalidateControl(ribbonControl.Id(), ribbonControl.Group());
                return; // not change if no more than 1 shape is selected
            }
            if (!IsOptionRibbonButton(ribbonControl)) {
                _alignRelativeFlag = _alignRelativeFlag switch {
                    ArrangementHelper.AlignRelativeFlag.RelativeToObjects => ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject,
                    ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject => ArrangementHelper.AlignRelativeFlag.RelativeToSlide,
                    ArrangementHelper.AlignRelativeFlag.RelativeToSlide => ArrangementHelper.AlignRelativeFlag.RelativeToObjects,
                    _ => ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject
                };
            } else {
                _alignRelativeFlag = ribbonControl.Id() switch {
                    btnAlignRelative_ToObjects => ArrangementHelper.AlignRelativeFlag.RelativeToObjects,
                    btnAlignRelative_ToFirstObject => ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject,
                    btnAlignRelative_ToSlide => ArrangementHelper.AlignRelativeFlag.RelativeToSlide,
                    _ => ArrangementHelper.AlignRelativeFlag.RelativeToObjects
                };
            }
            ArrangementHelper.ExecuteMsoByAlignRelative(_alignRelativeFlag);
            _ribbon?.InvalidateControl(btnAlignRelative, grpArrange);
            _ribbon?.InvalidateControl(btnDistributeHorizontal, grpArrange);
            _ribbon?.InvalidateControl(btnDistributeVertical, grpArrange);
            _ribbon?.InvalidateControl(btnAlignRelative_ToObjects, mnuArrangement);
            _ribbon?.InvalidateControl(btnAlignRelative_ToFirstObject, mnuArrangement);
            _ribbon?.InvalidateControl(btnAlignRelative_ToSlide, mnuArrangement);
            _ribbon?.InvalidateControl(btnAlignRelative_ToObjects, "grpAlignment");
            _ribbon?.InvalidateControl(btnAlignRelative_ToFirstObject, "grpAlignment");
            _ribbon?.InvalidateControl(btnAlignRelative_ToSlide, "grpAlignment");
        }

        public bool BtnAlignRelative_GetPressed(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                return false;
            }
            var shapeRange = GetShapeRange();
            if (shapeRange?.Count == 1) {
                return ribbonControl.Id() == btnAlignRelative_ToSlide; // when single shape is selected
            }
            return ribbonControl.Id() switch {
                btnAlignRelative_ToObjects => _alignRelativeFlag == ArrangementHelper.AlignRelativeFlag.RelativeToObjects,
                btnAlignRelative_ToFirstObject => _alignRelativeFlag == ArrangementHelper.AlignRelativeFlag.RelativeToFirstObject,
                btnAlignRelative_ToSlide => _alignRelativeFlag == ArrangementHelper.AlignRelativeFlag.RelativeToSlide,
                _ => false
            };
        }

        public string BtnAlignRelative_GetLabel(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
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
            return ribbonControl.Id() switch {
                btnAlignRelative_ToObjects => ArrangeRibbonResources.btnAlignRelative_ToObjects,
                btnAlignRelative_ToFirstObject => ArrangeRibbonResources.btnAlignRelative_ToFirstObject,
                btnAlignRelative_ToSlide => ArrangeRibbonResources.btnAlignRelative_ToSlide,
                _ => ArrangeRibbonResources.btnAlignRelative_ToObjects
            };
        }

        public System.Drawing.Image BtnAlignRelative_GetImage(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
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
            return ribbonControl.Id() switch {
                btnAlignRelative_ToObjects => Properties.Resources.AlignRelativeToObjects,
                btnAlignRelative_ToFirstObject => Properties.Resources.AlignRelativeToFirstObject,
                btnAlignRelative_ToSlide => Properties.Resources.AlignRelativeToSlide,
                _ => Properties.Resources.AlignRelativeToObjects
            };
        }

        public void BtnScale_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.ScaleSizeCmd? cmd = ribbonControl.Id() switch {
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
            if (!IsOptionRibbonButton(ribbonControl)) {
                _scaleFromFlag = _scaleFromFlag switch {
                    Office.MsoScaleFrom.msoScaleFromTopLeft => Office.MsoScaleFrom.msoScaleFromMiddle,
                    Office.MsoScaleFrom.msoScaleFromMiddle => Office.MsoScaleFrom.msoScaleFromBottomRight,
                    Office.MsoScaleFrom.msoScaleFromBottomRight => Office.MsoScaleFrom.msoScaleFromTopLeft,
                    _ => Office.MsoScaleFrom.msoScaleFromTopLeft
                };
            } else {
                _scaleFromFlag = ribbonControl.Id() switch {
                    btnScaleAnchor_FromTopLeft => Office.MsoScaleFrom.msoScaleFromTopLeft,
                    btnScaleAnchor_FromMiddle => Office.MsoScaleFrom.msoScaleFromMiddle,
                    btnScaleAnchor_FromBottomRight => Office.MsoScaleFrom.msoScaleFromBottomRight,
                    _ => _scaleFromFlag
                };
            }
            _ribbon?.InvalidateControl(btnScaleAnchor, grpArrange);
            _ribbon?.InvalidateControl(btnScaleAnchor, _sizeAndPositionGroups);
            _ribbon?.InvalidateControl(btnScaleAnchor_FromTopLeft, mnuArrangement);
            _ribbon?.InvalidateControl(btnScaleAnchor_FromMiddle, mnuArrangement);
            _ribbon?.InvalidateControl(btnScaleAnchor_FromBottomRight, mnuArrangement);
            _ribbon?.InvalidateControl(btnScaleAnchor_FromTopLeft, "grpSizeAndSnap");
            _ribbon?.InvalidateControl(btnScaleAnchor_FromMiddle, "grpSizeAndSnap");
            _ribbon?.InvalidateControl(btnScaleAnchor_FromBottomRight, "grpSizeAndSnap");
        }

        public bool BtnScaleAnchor_GetPressed(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                return false;
            }
            return ribbonControl.Id() switch {
                btnScaleAnchor_FromTopLeft => _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromTopLeft,
                btnScaleAnchor_FromMiddle => _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle,
                btnScaleAnchor_FromBottomRight => _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromBottomRight,
                _ => false
            };
        }

        public string BtnScaleAnchor_GetLabel(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                return _scaleFromFlag switch {
                    Office.MsoScaleFrom.msoScaleFromTopLeft => ArrangeRibbonResources.btnScaleAnchor_TopLeft,
                    Office.MsoScaleFrom.msoScaleFromMiddle => ArrangeRibbonResources.btnScaleAnchor_Middle,
                    Office.MsoScaleFrom.msoScaleFromBottomRight => ArrangeRibbonResources.btnScaleAnchor_BottomRight,
                    _ => ArrangeRibbonResources.btnScaleAnchor_TopLeft
                };
            }
            return ribbonControl.Id() switch {
                btnScaleAnchor_FromTopLeft => ArrangeRibbonResources.btnScaleAnchor_TopLeft,
                btnScaleAnchor_FromMiddle => ArrangeRibbonResources.btnScaleAnchor_Middle,
                btnScaleAnchor_FromBottomRight => ArrangeRibbonResources.btnScaleAnchor_BottomRight,
                _ => ArrangeRibbonResources.btnScaleAnchor_TopLeft
            };
        }

        public System.Drawing.Image BtnScaleAnchor_GetImage(Office.IRibbonControl ribbonControl) {
            if (!IsOptionRibbonButton(ribbonControl)) {
                return _scaleFromFlag switch {
                    Office.MsoScaleFrom.msoScaleFromTopLeft => Properties.Resources.ScaleFromTopLeft,
                    Office.MsoScaleFrom.msoScaleFromMiddle => Properties.Resources.ScaleFromMiddle,
                    Office.MsoScaleFrom.msoScaleFromBottomRight => Properties.Resources.ScaleFromBottomRight,
                    _ => Properties.Resources.ScaleFromTopLeft
                };
            }
            return ribbonControl.Id() switch {
                btnScaleAnchor_FromTopLeft => Properties.Resources.ScaleFromTopLeft,
                btnScaleAnchor_FromMiddle => Properties.Resources.ScaleFromMiddle,
                btnScaleAnchor_FromBottomRight => Properties.Resources.ScaleFromBottomRight,
                _ => Properties.Resources.ScaleFromTopLeft
            };
        }

        public void BtnExtend_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ArrangementHelper.ExtendSizeCmd? cmd = ribbonControl.Id() switch {
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
            ArrangementHelper.SnapCmd? cmd = ribbonControl.Id() switch {
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
            Office.MsoZOrderCmd? cmd = ribbonControl.Id() switch {
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
            ArrangementHelper.RotateCmd? cmd = ribbonControl.Id() switch {
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
            Office.MsoFlipCmd? cmd = ribbonControl.Id() switch {
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
            ArrangementHelper.GroupCmd? cmd = ribbonControl.Id() switch {
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

        public void BtnSizeAndPosition_Click(Office.IRibbonControl _) {
            var mso = "ObjectSizeAndPositionDialog";
            if (Globals.ThisAddIn.Application.CommandBars.GetEnabledMso(mso)) {
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso(mso);
            }
        }

        public void BtnAutofit_Click(Office.IRibbonControl ribbonControl, bool _) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            TextboxHelper.TextboxStatusCmd? cmd = ribbonControl.Id() switch {
                btnAutofitOff => TextboxHelper.TextboxStatusCmd.AutofitOff,
                btnAutoShrinkText => TextboxHelper.TextboxStatusCmd.AutoShrinkText,
                btnAutoResizeShape => TextboxHelper.TextboxStatusCmd.AutoResizeShape,
                _ => null
            };
            TextboxHelper.ChangeAutofitStatus(textFrame, cmd, () => {
                _ribbon?.InvalidateControl(btnAutofitOff, grpTextbox);
                _ribbon?.InvalidateControl(btnAutoShrinkText, grpTextbox);
                _ribbon?.InvalidateControl(btnAutoResizeShape, grpTextbox);
            });
        }

        public bool BtnAutofit_GetPressed(Office.IRibbonControl ribbonControl) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return false;
            }
            TextboxHelper.TextboxStatusCmd? cmd = ribbonControl.Id() switch {
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
                _ribbon?.InvalidateControl(btnWrapText, grpTextbox);
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
            TextboxHelper.ResetMarginCmd? cmd = ribbonControl.Id() switch {
                btnResetHorizontalMargin => TextboxHelper.ResetMarginCmd.Horizontal,
                btnResetVerticalMargin => TextboxHelper.ResetMarginCmd.Vertical,
                _ => null
            };
            TextboxHelper.ResetMargin(textFrame, cmd, () => {
                _ribbon?.InvalidateControl(edtMarginLeft, grpTextbox);
                _ribbon?.InvalidateControl(edtMarginRight, grpTextbox);
                _ribbon?.InvalidateControl(edtMarginTop, grpTextbox);
                _ribbon?.InvalidateControl(edtMarginBottom, grpTextbox);
            });
        }

        public void EdtMargin_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            TextboxHelper.MarginKind? kind = ribbonControl.Id() switch {
                edtMarginLeft => TextboxHelper.MarginKind.Left,
                edtMarginRight => TextboxHelper.MarginKind.Right,
                edtMarginTop => TextboxHelper.MarginKind.Top,
                edtMarginBottom => TextboxHelper.MarginKind.Bottom,
                _ => null
            };
            TextboxHelper.ChangeMarginOfString(textFrame, kind, text, () => {
                _ribbon?.InvalidateControl(ribbonControl.Id(), ribbonControl.Group());
            });
        }

        public string EdtMargin_GetText(Office.IRibbonControl ribbonControl) {
            var (textFrame, _) = GetTextFrame();
            if (textFrame == null) {
                return "";
            }
            TextboxHelper.MarginKind? kind = ribbonControl.Id() switch {
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
                _ribbon?.InvalidateControl(btnLockAspectRatio, grpTextbox);
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
            SizeAndPositionHelper.CopyAndPasteCmd? cmd = ribbonControl.Id() switch {
                btnCopySize => SizeAndPositionHelper.CopyAndPasteCmd.Copy,
                btnPasteSize => SizeAndPositionHelper.CopyAndPasteCmd.Paste,
                _ => null
            };
            SizeAndPositionHelper.CopyAndPasteSize(shapeRange, cmd, _scaleFromFlag, () => {
                _ribbon?.InvalidateControl(btnPasteSize, _sizeAndPositionGroups);
            });
        }

        public void EdtPosition_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.PositionKind? kind = ribbonControl.Id() switch {
                edtPositionX => SizeAndPositionHelper.PositionKind.X,
                edtPositionY => SizeAndPositionHelper.PositionKind.Y,
                _ => null
            };
            SizeAndPositionHelper.ChangePositionOfString(shapeRange, kind, text, () => {
                _ribbon?.InvalidateControl(ribbonControl.Id(), ribbonControl.Group());
            });
        }

        public string EdtPosition_GetText(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return "";
            }
            SizeAndPositionHelper.PositionKind? kind = ribbonControl.Id() switch {
                edtPositionX => SizeAndPositionHelper.PositionKind.X,
                edtPositionY => SizeAndPositionHelper.PositionKind.Y,
                _ => null
            };
            return SizeAndPositionHelper.GetPositionOfString(shapeRange, kind).Item1;
        }

        public void BtnCopyAndPastePosition_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.CopyAndPasteCmd? cmd = ribbonControl.Id() switch {
                btnCopyPosition => SizeAndPositionHelper.CopyAndPasteCmd.Copy,
                btnPastePosition => SizeAndPositionHelper.CopyAndPasteCmd.Paste,
                _ => null
            };
            SizeAndPositionHelper.CopyAndPastePosition(shapeRange, cmd, () => {
                _ribbon?.InvalidateControl(btnPastePosition, _sizeAndPositionGroups);
                _ribbon?.InvalidateControl(edtPositionX, _sizeAndPositionGroups);
                _ribbon?.InvalidateControl(edtPositionY, _sizeAndPositionGroups);
            });
        }

        public void BtnResetMediaSize_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.ResetMediaSize(shapeRange, _scaleFromFlag, () => {
                var groups = new[] { grpPictureSizeAndPosition, grpVideoSizeAndPosition, grpAudioSizeAndPosition };
                _ribbon?.InvalidateControl(edtPositionX, groups);
                _ribbon?.InvalidateControl(edtPositionY, groups);
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
            _ribbon?.InvalidateControl(chkReserveOriginalSize, grpReplacePicture);
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
            _ribbon?.InvalidateControl(chkReplaceToMiddle, grpReplacePicture);
        }

        public bool ChkReplaceToMiddle_GetPressed(Office.IRibbonControl _) {
            return (_replacePictureFlag & ReplacePictureHelper.ReplacePictureFlag.ReplaceToMiddle) != 0;
        }

        public void BtnReplacePicture_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            ReplacePictureHelper.ReplacePictureCmd? cmd = ribbonControl.Id() switch {
                btnReplaceWithClipboard => ReplacePictureHelper.ReplacePictureCmd.WithClipboard,
                btnReplaceWithFile => ReplacePictureHelper.ReplacePictureCmd.WithFile,
                _ => null
            };
            ReplacePictureHelper.ReplacePicture(shapeRange, cmd, _replacePictureFlag, () => _ribbon?.Invalidate());
        }

    }

}
