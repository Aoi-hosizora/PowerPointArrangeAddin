using System.Collections.Generic;
using ppt_arrange_addin.Helper;
using Forms = System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ppt_arrange_addin {

    public partial class ArrangeRibbon {

        private Office.IRibbonUI _ribbon;

        public void Ribbon_Load(Office.IRibbonUI ribbonUi) {
            _ribbon = ribbonUi;
        }

        private delegate bool AvailabilityRule(PowerPoint.ShapeRange shapeRange, int shapesCount, bool hasTextFrame);
        private Dictionary<string, AvailabilityRule> _availabilityRules;

        private void InitializeAvailabilityRules() {
            _availabilityRules = new Dictionary<string, AvailabilityRule> {
                // grpArrange
                { btnAlignLeft, (_, cnt, _) => cnt >= 1 },
                { btnAlignCenter, (_, cnt, _) => cnt >= 1 },
                { btnAlignRight, (_, cnt, _) => cnt >= 1 },
                { btnAlignTop, (_, cnt, _) => cnt >= 1 },
                { btnAlignMiddle, (_, cnt, _) => cnt >= 1 },
                { btnAlignBottom, (_, cnt, _) => cnt >= 1 },
                { btnDistributeHorizontal, (_, cnt, _) => cnt is 1 or >= 3 },
                { btnDistributeVertical, (_, cnt, _) => cnt is 1 or >= 3 },
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
                { btnPastePicturePosition, (_, cnt, _) => cnt >= 1 && SizeAndPositionHelper.IsValidCopiedPositionValue() }
            };
        }

        public bool GetEnabled(Office.IRibbonControl ribbonControl) {
            var selection = SelectionGetter.GetSelection(onlyShapeRange: false);
            var shapesCount = selection.ShapeRange?.Count ?? 0;
            var hasTextFrame = selection.TextFrame != null;
            _availabilityRules.TryGetValue(ribbonControl.Id, out var checker);
            return checker?.Invoke(selection.ShapeRange, shapesCount, hasTextFrame) ?? true;
        }

        public void InvalidateRibbon(bool onlyForDrag = false) {
            if (_ribbon == null) {
                return;
            }
            if (!onlyForDrag) {
                _ribbon.Invalidate();
            } else {
                // currently callback that only for dragging to change the position is unavailable
                _ribbon.InvalidateControl(edtShapePositionX);
                _ribbon.InvalidateControl(edtShapePositionY);
            }
        }

        public bool GetGroupVisible(Office.IRibbonControl ribbonControl) {
            return ribbonControl.Id switch {
                grpWordArt => AddInSetting.Instance.ShowWordArtGroup,
                grpArrange => true,
                grpTextbox => AddInSetting.Instance.ShowShapeTextboxGroup,
                grpShapeSizeAndPosition => AddInSetting.Instance.ShowShapeSizeAndPositionGroup,
                grpReplacePicture => AddInSetting.Instance.ShowReplacePictureGroup,
                grpPictureSizeAndPosition => AddInSetting.Instance.ShowPictureSizeAndPositionGroup,
                _ => true
            };
        }

        private PowerPoint.ShapeRange GetShapeRange(int mustMoreThanOrEqualTo = 1) {
            var selection = SelectionGetter.GetSelection(onlyShapeRange: true);
            var shapeRange = selection.ShapeRange;
            if (shapeRange == null || shapeRange.Count < mustMoreThanOrEqualTo) {
                return null;
            }
            return shapeRange;
        }

        private (PowerPoint.TextFrame, PowerPoint.TextFrame2) GetTextFrame() {
            var selection = SelectionGetter.GetSelection(onlyShapeRange: false);
            return (selection.TextFrame, selection.TextFrame2);
        }

        private void StartNewUndoEntry() {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
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
            ArrangementHelper.Align(shapeRange, cmd);
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
            ArrangementHelper.Distribute(shapeRange, cmd);
        }

        // This flag is used by scale and size related callbacks, that is BtnScale_Click,
        // BtnCopyAndPasteSize_Click and BtnResetPictureSize_Click.
        private Office.MsoScaleFrom _scaleFromFlag = Office.MsoScaleFrom.msoScaleFromMiddle;

        public void BtnScaleAnchor_Click(Office.IRibbonControl _) {
            _scaleFromFlag = _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle
                ? Office.MsoScaleFrom.msoScaleFromTopLeft
                : Office.MsoScaleFrom.msoScaleFromMiddle;
            _ribbon.InvalidateControl(btnScaleAnchor);
            _ribbon.InvalidateControl(btnShapeScaleAnchor);
            _ribbon.InvalidateControl(btnPictureScaleAnchor);
        }

        public string BtnScaleAnchor_GetLabel(Office.IRibbonControl _) {
            return _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle
                ? ArrangeRibbonResources.btnScaleAnchor_Middle
                : ArrangeRibbonResources.btnScaleAnchor_TopLeft;
        }

        public System.Drawing.Image BtnScaleAnchor_GetImage(Office.IRibbonControl _) {
            return _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle
                ? Properties.Resources.ScaleFromMiddle
                : Properties.Resources.ScaleFromTopLeft;
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
                btnSnapLeft => ArrangementHelper.SnapCmd.SnapToLeft,
                btnSnapRight => ArrangementHelper.SnapCmd.SnapToRight,
                btnSnapTop => ArrangementHelper.SnapCmd.SnapToTop,
                btnSnapBottom => ArrangementHelper.SnapCmd.SnapToBottom,
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
            ArrangementHelper.Group(shapeRange, cmd, () => InvalidateRibbon());
        }

        public string MnuArrangement_GetContent(Office.IRibbonControl _) {
            return GetResourceText("ppt_arrange_addin.ArrangeRibbon.ArrangeMenu.xml");
        }

        public void BtnAddInSetting_Click(Office.IRibbonControl _) {
            var dlg = new SettingDialog();
            var result = dlg.ShowDialog();
            if (result == Forms.DialogResult.OK) {
                _ribbon.Invalidate();
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
                _ribbon.InvalidateControl(btnAutofitOff);
                _ribbon.InvalidateControl(btnAutoShrinkText);
                _ribbon.InvalidateControl(btnAutoResizeShape);
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

        public void BtnWrapText_Click(Office.IRibbonControl ribbonControl, bool _) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return;
            }
            var cmd = TextboxHelper.TextboxStatusCmd.WrapTextOnOff;
            TextboxHelper.ChangeAutofitStatus(textFrame, cmd, () => {
                _ribbon.InvalidateControl(btnWrapText);
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
                _ribbon.InvalidateControl(ribbonControl.Id);
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
                _ribbon.InvalidateControl(edtMarginLeft);
                _ribbon.InvalidateControl(edtMarginRight);
                _ribbon.InvalidateControl(edtMarginTop);
                _ribbon.InvalidateControl(edtMarginBottom);
            });
        }

        public void BtnLockAspectRatio_Click(Office.IRibbonControl ribbonControl, bool pressed) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            var cmd = SizeAndPositionHelper.LockAspectRatioCmd.Toggle;
            SizeAndPositionHelper.ToggleLockAspectRatio(shapeRange, cmd, () => {
                _ribbon.InvalidateControl(ribbonControl.Id);
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
                btnCopyShapeSize or btnCopyPictureSize => SizeAndPositionHelper.CopyAndPasteCmd.Copy,
                btnPasteShapeSize or btnPastePictureSize => SizeAndPositionHelper.CopyAndPasteCmd.Paste,
                _ => null
            };
            SizeAndPositionHelper.CopyAndPasteSize(shapeRange, cmd, _scaleFromFlag, () => {
                _ribbon.InvalidateControl(btnPasteShapeSize);
                _ribbon.InvalidateControl(btnPastePictureSize);
            });
        }

        public void EdtPosition_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.PositionKind? kind = ribbonControl.Id switch {
                edtShapePositionX or edtPicturePositionX => SizeAndPositionHelper.PositionKind.X,
                edtShapePositionY or edtPicturePositionY => SizeAndPositionHelper.PositionKind.Y,
                _ => null
            };
            SizeAndPositionHelper.ChangePositionOfString(shapeRange, kind, text, () => {
                _ribbon.InvalidateControl(ribbonControl.Id);
            });
        }

        public string EdtPosition_GetText(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return "";
            }
            SizeAndPositionHelper.PositionKind? kind = ribbonControl.Id switch {
                edtShapePositionX or edtPicturePositionX => SizeAndPositionHelper.PositionKind.X,
                edtShapePositionY or edtPicturePositionY => SizeAndPositionHelper.PositionKind.Y,
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
                btnCopyShapePosition or btnCopyPicturePosition => SizeAndPositionHelper.CopyAndPasteCmd.Copy,
                btnPasteShapePosition or btnPastePicturePosition => SizeAndPositionHelper.CopyAndPasteCmd.Paste,
                _ => null
            };
            SizeAndPositionHelper.CopyAndPastePosition(shapeRange, cmd, () => {
                _ribbon.InvalidateControl(btnPasteShapePosition);
                _ribbon.InvalidateControl(btnPastePicturePosition);
                _ribbon.InvalidateControl(edtShapePositionX);
                _ribbon.InvalidateControl(edtShapePositionY);
                _ribbon.InvalidateControl(edtPicturePositionX);
                _ribbon.InvalidateControl(edtPicturePositionY);
            });
        }

        public void BtnResetPictureSize_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            SizeAndPositionHelper.ResetPictureSize(shapeRange, _scaleFromFlag, () => {
                _ribbon.InvalidateControl(edtPicturePositionX);
                _ribbon.InvalidateControl(edtPicturePositionY);
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
            _ribbon.InvalidateControl(ribbonControl.Id);
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
            _ribbon.InvalidateControl(ribbonControl.Id);
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
            ReplacePictureHelper.ReplacePicture(shapeRange, cmd, _replacePictureFlag, () => InvalidateRibbon());
        }

    }

}
