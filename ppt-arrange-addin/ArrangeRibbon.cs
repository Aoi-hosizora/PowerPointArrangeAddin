using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private struct Selection {
            public PowerPoint.ShapeRange ShapeRange { get; init; }
            public PowerPoint.Shape TextShape { get; init; }
            public PowerPoint.TextRange TextRange { get; init; }
            public PowerPoint.TextFrame TextFrame { get; init; }
            public PowerPoint.TextFrame2 TextFrame2 { get; init; }
        }

        private Selection GetSelection(bool onlyShapeRange) {
            // 1. application
            PowerPoint.Selection selection = null;
            try {
                var application = Globals.ThisAddIn.Application;
                if (application.Windows.Count > 0 /* GetForegroundWindow().ToInt32() == application.HWND */) {
                    selection = application.ActiveWindow.Selection;
                }
            } catch (Exception) { /* ignored */ }
            if (selection == null) {
                return new Selection();
            }

            // 2. shape range
            PowerPoint.ShapeRange shapeRange = null;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) {
                shapeRange = selection.ShapeRange;
            } else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText) {
                try {
                    shapeRange = selection.ShapeRange;
                } catch (Exception) { /* ignored */ }
            }
            if (onlyShapeRange) {
                return new Selection { ShapeRange = shapeRange };
            }

            // 3. text range
            PowerPoint.TextRange textRange = null;
            PowerPoint.TextFrame textFrame = null;
            PowerPoint.Shape textShape = null;
            PowerPoint.TextFrame2 textFrame2 = null;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText) {
                textRange = selection.TextRange;
                if (textRange.Parent is PowerPoint.TextFrame frame) {
                    textFrame = frame;
                    if (textFrame.Parent is PowerPoint.Shape shape) {
                        textShape = shape;
                        textFrame2 = shape.TextFrame2;
                    }
                }
            } else if (shapeRange != null && shapeRange.HasTextFrame != Office.MsoTriState.msoFalse) {
                textFrame = shapeRange.TextFrame;
                textRange = textFrame.TextRange;
                textFrame2 = shapeRange.TextFrame2;
            }

            // 4. return selection
            return new Selection {
                ShapeRange = shapeRange,
                TextRange = textRange,
                TextShape = textShape,
                TextFrame = textFrame,
                TextFrame2 = textFrame2
            };
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
                { btnScalePosition, (_, _, _) => true },
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
                { btnUngroup, (shapeRange, cnt, _) => cnt >= 1 && IsUngroupable(shapeRange) },
                { mnuArrangement, (_, _, _) => true },
                { btnAddInSetting, (_, _, _) => true },
                // grpTextbox
                { btnAutofitOff, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnAutofitText, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnAutofitShape, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnWrapText, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginLeft, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginRight, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginTop, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { edtMarginBottom, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnResetMarginHorizontal, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                { btnResetMarginVertical, (_, cnt, hasTextFrame) => cnt >= 1 && hasTextFrame },
                // grpShapeSizeAndPosition
                { mnuShapeArrangement, (_, _, _) => true },
                { btnLockShapeAspectRatio, (_, cnt, _) => cnt >= 1 },
                { btnShapeScalePosition, (_, _, _) => true },
                { btnCopyShapeSize, (_, cnt, _) => cnt == 1 },
                { btnPasteShapeSize, (_, cnt, _) => cnt >= 1 && IsValidCopiedSizeValue() },
                { edtShapePositionX, (_, cnt, _) => cnt >= 1 },
                { edtShapePositionY, (_, cnt, _) => cnt >= 1 },
                { btnCopyShapePosition, (_, cnt, _) => cnt == 1 },
                { btnPasteShapePosition, (_, cnt, _) => cnt >= 1 && IsValidCopiedPositionValue() },
                // grpReplacePicture
                { btnReplaceWithClipboard, (_, cnt, _) => cnt >= 1 },
                { btnReplaceWithFile, (_, cnt, _) => cnt >= 1 },
                { cbxReserveOriginalSize, (_, _, _) => true },
                { cbxReplaceToMiddle, (_, _, _) => true },
                // grpPictureSizeAndPosition
                { mnuPictureArrangement, (_, _, _) => true },
                { btnResetPictureSize, (_,cnt, _) => cnt >= 1 },
                { btnLockPictureAspectRatio, (_, cnt, _) => cnt >= 1 },
                { btnPictureScalePosition, (_, _, _) => true },
                { btnCopyPictureSize, (_, cnt, _) => cnt == 1 },
                { btnPastePictureSize, (_, cnt, _) => cnt >= 1 && IsValidCopiedSizeValue() },
                { edtPicturePositionX, (_, cnt, _) => cnt >= 1 },
                { edtPicturePositionY, (_, cnt, _) => cnt >= 1 },
                { btnCopyPicturePosition, (_, cnt, _) => cnt == 1 },
                { btnPastePicturePosition, (_, cnt, _) => cnt >= 1 && IsValidCopiedPositionValue() }
            };
        }

        public bool GetEnabled(Office.IRibbonControl ribbonControl) {
            var selection = GetSelection(onlyShapeRange: false);
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
            var selection = GetSelection(onlyShapeRange: true);
            var shapeRange = selection.ShapeRange;
            if (shapeRange == null || shapeRange.Count < mustMoreThanOrEqualTo) {
                return null;
            }
            return shapeRange;
        }

        private (PowerPoint.TextFrame, PowerPoint.TextFrame2) GetTextFrame() {
            var selection = GetSelection(onlyShapeRange: false);
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

        private Office.MsoScaleFrom _scaleFromFlag = Office.MsoScaleFrom.msoScaleFromMiddle; // used by BtnScale_Click and BtnCopyAndPasteSize_Click

        public void BtnScalePosition_Click(Office.IRibbonControl _) {
            _scaleFromFlag = _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle
                ? Office.MsoScaleFrom.msoScaleFromTopLeft
                : Office.MsoScaleFrom.msoScaleFromMiddle;
            _ribbon.InvalidateControl(btnScalePosition);
            _ribbon.InvalidateControl(btnShapeScalePosition);
            _ribbon.InvalidateControl(btnPictureScalePosition);
        }

        public string GetBtnScalePositionLabel(Office.IRibbonControl _) {
            return _scaleFromFlag == Office.MsoScaleFrom.msoScaleFromMiddle
                ? ArrangeRibbonResources.btnScalePosition_Middle
                : ArrangeRibbonResources.btnScalePosition_TopLeft;
        }

        public System.Drawing.Image GetBtnScalePositionImage(Office.IRibbonControl _) {
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

        private bool IsUngroupable(PowerPoint.ShapeRange shapeRange) {
            return ArrangementHelper.IsUngroupable(shapeRange);
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

        public string GetMnuArrangementContent(Office.IRibbonControl _) {
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
                btnAutofitText => TextboxHelper.TextboxStatusCmd.AutofitText,
                btnAutofitShape => TextboxHelper.TextboxStatusCmd.AutofitShape,
                _ => null
            };
            TextboxHelper.ChangeAutofitStatus(textFrame, cmd, () => {
                _ribbon.InvalidateControl(btnAutofitOff);
                _ribbon.InvalidateControl(btnAutofitText);
                _ribbon.InvalidateControl(btnAutofitShape);
            });
        }

        public bool GetBtnAutofitPressed(Office.IRibbonControl ribbonControl) {
            var (_, textFrame) = GetTextFrame();
            if (textFrame == null) {
                return false;
            }
            TextboxHelper.TextboxStatusCmd? cmd = ribbonControl.Id switch {
                btnAutofitOff => TextboxHelper.TextboxStatusCmd.AutofitOff,
                btnAutofitText => TextboxHelper.TextboxStatusCmd.AutofitText,
                btnAutofitShape => TextboxHelper.TextboxStatusCmd.AutofitShape,
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

        public bool GetBtnWrapTextPressed(Office.IRibbonControl _) {
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

        public string GetEdtMarginText(Office.IRibbonControl ribbonControl) {
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
                btnResetMarginHorizontal => TextboxHelper.ResetMarginCmd.Horizontal,
                btnResetMarginVertical => TextboxHelper.ResetMarginCmd.Vertical,
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

            StartNewUndoEntry();
            shapeRange.LockAspectRatio = shapeRange.LockAspectRatio != Office.MsoTriState.msoTrue
                ? Office.MsoTriState.msoTrue
                : Office.MsoTriState.msoFalse;
            _ribbon.InvalidateControl(ribbonControl.Id);
        }

        public bool GetBtnLockAspectRatio(Office.IRibbonControl _) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return false;
            }

            return shapeRange.LockAspectRatio == Office.MsoTriState.msoTrue;
        }

        private const float InvalidCopiedValue = -2147483648.0F; // for size and position
        private float _copiedSizeWPt = InvalidCopiedValue; // for shape and image
        private float _copiedSizeHPt = InvalidCopiedValue; // for shape and image

        private bool IsValidCopiedSizeValue() {
            return !_copiedSizeWPt.Equals(InvalidCopiedValue) && !_copiedSizeHPt.Equals(InvalidCopiedValue);
        }

        public void BtnCopyAndPasteSize_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            switch (ribbonControl.Id) {
            case btnCopyShapeSize:
            case btnCopyPictureSize:
                if (shapeRange.Count == 1) {
                    StartNewUndoEntry();
                    _copiedSizeWPt = shapeRange.Width;
                    _copiedSizeHPt = shapeRange.Height;
                    _ribbon.InvalidateControl(btnPasteShapeSize);
                    _ribbon.InvalidateControl(btnPastePictureSize);
                }
                break;
            case btnPasteShapeSize:
            case btnPastePictureSize:
                if (IsValidCopiedSizeValue()) {
                    StartNewUndoEntry();
                    foreach (var shape in shapeRange.OfType<PowerPoint.Shape>().ToArray()) {
                        var oldLockState = shape.LockAspectRatio;
                        shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                        var ratio = _copiedSizeWPt / shape.Width;
                        shape.ScaleWidth(ratio, Office.MsoTriState.msoFalse, _scaleFromFlag);
                        ratio = _copiedSizeHPt / shape.Height;
                        shape.ScaleHeight(ratio, Office.MsoTriState.msoFalse, _scaleFromFlag);
                        shape.LockAspectRatio = oldLockState;
                    }
                }
                break;
            }
        }

        public void EdtPosition_TextChanged(Office.IRibbonControl ribbonControl, string text) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            text = text.Replace("cm", "").Trim();
            if (text.Length == 0) {
                text = "0";
            }

            StartNewUndoEntry();
            if (float.TryParse(text, out var input)) {
                var pt = UnitConverter.CmToPt(input);
                switch (ribbonControl.Id) {
                case edtShapePositionX:
                case edtPicturePositionX:
                    shapeRange.Left = pt;
                    break;
                case edtShapePositionY:
                case edtPicturePositionY:
                    shapeRange.Top = pt;
                    break;
                }
            }

            _ribbon.InvalidateControl(ribbonControl.Id);
        }

        public string GetEdtPositionText(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return "";
            }

            var pt = ribbonControl.Id switch {
                edtShapePositionX or edtPicturePositionX => shapeRange.Left,
                edtShapePositionY or edtPicturePositionY => shapeRange.Top,
                _ => -1
            };

            return pt < 0
                ? ""
                : $"{Math.Round(UnitConverter.PtToCm(pt), 2)} cm";
        }

        private float _copiedPositionXPt = InvalidCopiedValue; // for shape and image
        private float _copiedPositionYPt = InvalidCopiedValue; // for shape and image

        private bool IsValidCopiedPositionValue() {
            return !_copiedPositionXPt.Equals(InvalidCopiedValue) && !_copiedPositionYPt.Equals(InvalidCopiedValue);
        }

        public void BtnCopyAndPastePosition_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            switch (ribbonControl.Id) {
            case btnCopyShapePosition:
            case btnCopyPicturePosition:
                if (shapeRange.Count == 1) {
                    StartNewUndoEntry();
                    _copiedPositionXPt = shapeRange.Left;
                    _copiedPositionYPt = shapeRange.Top;
                    _ribbon.InvalidateControl(btnPasteShapePosition);
                    _ribbon.InvalidateControl(btnPastePicturePosition);
                }
                break;
            case btnPasteShapePosition:
            case btnPastePicturePosition:
                if (IsValidCopiedPositionValue()) {
                    StartNewUndoEntry();
                    shapeRange.Left = _copiedPositionXPt;
                    shapeRange.Top = _copiedPositionYPt;
                    _ribbon.InvalidateControl(edtShapePositionX);
                    _ribbon.InvalidateControl(edtShapePositionY);
                    _ribbon.InvalidateControl(edtPicturePositionX);
                    _ribbon.InvalidateControl(edtPicturePositionY);
                }
                break;
            }
        }

        private bool _reserveOriginalSize = true; // used by replacing picture
        private bool _replaceToMiddle = true; // used by replacing picture

        public void CbxReserveOriginalSize_Click(Office.IRibbonControl ribbonControl, bool _) {
            _reserveOriginalSize = !_reserveOriginalSize;
            _ribbon.InvalidateControl(ribbonControl.Id);
        }

        public bool GetCbxReserveOriginalSize(Office.IRibbonControl _) {
            return _reserveOriginalSize;
        }

        public void CbxReplaceToMiddle_Click(Office.IRibbonControl ribbonControl, bool _) {
            _replaceToMiddle = !_replaceToMiddle;
            _ribbon.InvalidateControl(ribbonControl.Id);
        }

        public bool GetCbxReplaceToMiddlePressed(Office.IRibbonControl _) {
            return _replaceToMiddle;
        }

        public void BtnReplacePicture_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }
            var pictures = shapeRange.OfType<PowerPoint.Shape>().Where(shape => shape.Type == Office.MsoShapeType.msoPicture).ToArray();
            if (pictures.Length == 0) {
                return;
            }

            PowerPoint.Shapes slideShapes = null;
            if (shapeRange.Parent is PowerPoint.Slide slide) {
                slideShapes = slide.Shapes;
            }
            if (slideShapes == null) {
                return;
            }

            var (path, needCleanup) = ("", false);
            switch (ribbonControl.Id) {
            case btnReplaceWithClipboard:
                var image = Forms.Clipboard.GetImage();
                if (image == null) {
                    Forms.MessageBox.Show(ArrangeRibbonResources.dlgNoPictureInClipboard, ArrangeRibbonResources.dlgReplacePicture,
                        Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error);
                    return;
                }
                path = Path.GetTempFileName();
                needCleanup = true;
                try {
                    image.Save(path, ImageFormat.Png);
                } catch (Exception) {
                    return;
                }
                break;
            case btnReplaceWithFile:
                var dlg = Globals.ThisAddIn.Application.FileDialog[Office.MsoFileDialogType.msoFileDialogFilePicker];
                dlg.Title = ArrangeRibbonResources.dlgSelectPictureToReplace;
                dlg.AllowMultiSelect = false;
                dlg.Filters.Add("Image files", "*.jpg; *.jpeg; *.png; *.bmp");
                dlg.Filters.Add("All files", "*.*");
                if (dlg.Show() != -1 && dlg.SelectedItems.Count != 0) {
                    return;
                }
                path = dlg.SelectedItems.Item(1);
                break;
            }
            if (path.Length == 0) {
                return;
            }

            StartNewUndoEntry();
            var newShapes = new List<PowerPoint.Shape>();
            foreach (var shape in pictures) {
                try {
                    var (toLink, toSaveWith) = (Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
                    var newShape = slideShapes.AddPicture(path, toLink, toSaveWith, shape.Left, shape.Top);
                    newShape.LockAspectRatio = shape.LockAspectRatio;
                    // TODO apply old format
                    var (oldWidth, oldHeight) = (shape.Width, shape.Height);
                    var (oldLeft, oldTop) = (shape.Left, shape.Top);
                    var (newWidth, newHeight) = (newShape.Width, newShape.Height);
                    if (_reserveOriginalSize) {
                        var widthHeightRate = newWidth / newHeight;
                        if (oldHeight * widthHeightRate <= oldWidth) {
                            newHeight = oldHeight;
                            newWidth = oldHeight * widthHeightRate;
                        } else {
                            newWidth = oldWidth;
                            newHeight = oldWidth / widthHeightRate;
                        }
                        newShape.Width = newWidth;
                        newShape.Height = newHeight;
                    }
                    if (_replaceToMiddle) {
                        newShape.Left = oldLeft - (newWidth - oldWidth) / 2;
                        newShape.Top = oldTop - (newHeight - oldHeight) / 2;
                    }
                    newShapes.Add(newShape);
                    shape.Delete();
                } catch (Exception) {
                    // ignored
                }
            }

            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            foreach (var shape in newShapes) {
                shape.Select(Office.MsoTriState.msoFalse);
            }

            if (needCleanup) {
                try {
                    File.Delete(path);
                } catch (Exception) {
                    // ignored
                }
            }
        }

        public void BtnResetPictureSize_Click(Office.IRibbonControl ribbonControl) {
            var shapeRange = GetShapeRange();
            if (shapeRange == null) {
                return;
            }

            StartNewUndoEntry();
            shapeRange.ScaleWidth(1F, Office.MsoTriState.msoTrue); // scale from top right always
            shapeRange.ScaleHeight(1F, Office.MsoTriState.msoTrue);
        }

    }

}
