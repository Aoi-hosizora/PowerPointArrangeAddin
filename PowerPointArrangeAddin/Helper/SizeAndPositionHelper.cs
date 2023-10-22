using System;
using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

#nullable enable

namespace PowerPointArrangeAddin.Helper {

    public static class SizeAndPositionHelper {

        public static void ScaleSizeTo(this PowerPoint.Shape shape, float? width, float? height, Office.MsoScaleFrom? scaleFromFlag) {
            var oldLockSate = shape.LockAspectRatio;
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            scaleFromFlag ??= Office.MsoScaleFrom.msoScaleFromTopLeft;

            var (oldLeft, oldTop) = (shape.Left, shape.Top);
            var (oldWidth, oldHeight) = (shape.Width, shape.Height);
            if (width != null) shape.Width = width.Value;
            if (height != null) shape.Height = height.Value;
            var (newWidth, newHeight) = (shape.Width, shape.Height);

            switch (scaleFromFlag) {
            case Office.MsoScaleFrom.msoScaleFromTopLeft:
                break;
            case Office.MsoScaleFrom.msoScaleFromMiddle:
                shape.Left = oldLeft - (newWidth - oldWidth) / 2;
                shape.Top = oldTop - (newHeight - oldHeight) / 2;
                break;
            case Office.MsoScaleFrom.msoScaleFromBottomRight:
                shape.Left = oldLeft - (newWidth - oldWidth);
                shape.Top = oldTop - (newHeight - oldHeight);
                break;
            }

            shape.LockAspectRatio = oldLockSate;
        }

        public static void ScaleWidthTo(this PowerPoint.Shape shape, float width, Office.MsoScaleFrom scaleFromFlag) {
            shape.ScaleSizeTo(width, null, scaleFromFlag);
        }

        public static void ScaleHeightTo(this PowerPoint.Shape shape, float height, Office.MsoScaleFrom scaleFromFlag) {
            shape.ScaleSizeTo(null, height, scaleFromFlag);
        }

        public enum LockAspectRatioCmd {
            Lock,
            Unlock,
            Toggle
        }

        public static void ToggleLockAspectRatio(PowerPoint.ShapeRange? shapeRange, LockAspectRatioCmd? cmd, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (cmd == null) {
                return;
            }

            Office.MsoTriState? state = cmd! switch {
                LockAspectRatioCmd.Lock => Office.MsoTriState.msoTrue,
                LockAspectRatioCmd.Unlock => Office.MsoTriState.msoFalse,
                LockAspectRatioCmd.Toggle => shapeRange.LockAspectRatio != Office.MsoTriState.msoTrue
                    ? Office.MsoTriState.msoTrue
                    : Office.MsoTriState.msoFalse,
                _ => null
            };
            if (state == null) {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            shapeRange.LockAspectRatio = state.Value;
            uiInvalidator?.Invoke();
        }

        public static bool GetAspectRatioIsLocked(PowerPoint.ShapeRange? shapeRange) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return false;
            }
            return shapeRange.LockAspectRatio == Office.MsoTriState.msoTrue;
        }

        public enum PositionKind {
            X,
            Y
        }

        public static void ChangePositionOfString(PowerPoint.ShapeRange? shapeRange, PositionKind? positionKind, string? input, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (positionKind == null || input == null) {
                return;
            }
            var (valueInPt, ok) = UnitConverter.ParseStringToPtValue(input);
            if (!ok) {
                uiInvalidator?.Invoke(); // reset input
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (positionKind) {
            case PositionKind.X:
                shapeRange.Left = valueInPt;
                break;
            case PositionKind.Y:
                shapeRange.Top = valueInPt;
                break;
            }
            uiInvalidator?.Invoke();
        }

        public static (string, bool) GetPositionOfString(PowerPoint.ShapeRange? shapeRange, PositionKind? positionKind) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return ("", false);
            }
            if (positionKind == null) {
                return ("", false);
            }

            var valueInPt = positionKind! switch {
                PositionKind.X => shapeRange.Left,
                PositionKind.Y => shapeRange.Top,
                _ => -1e9F // if shapes has different position, the value will be "-2.147484E+09"
            };

            var text = "";
            if (valueInPt >= -1e9F) {
                text = UnitConverter.FormatPtValueToString(valueInPt);
            }
            return (text, true);
        }

        private const float InvalidCopiedValue = -2147483648.0F; // for size and position

        private static float _copiedSizeWPt = InvalidCopiedValue;
        private static float _copiedSizeHPt = InvalidCopiedValue;

        private static float _copiedPositionXPt = InvalidCopiedValue;
        private static float _copiedPositionYPt = InvalidCopiedValue;

        public static bool IsValidCopiedSizeValue() {
            return !_copiedSizeWPt.Equals(InvalidCopiedValue) && !_copiedSizeHPt.Equals(InvalidCopiedValue);
        }

        public static bool IsValidCopiedPositionValue() {
            return !_copiedPositionXPt.Equals(InvalidCopiedValue) && !_copiedPositionYPt.Equals(InvalidCopiedValue);
        }

        public enum CopyAndPasteCmd {
            Copy,
            Paste,
            Reset
        }

        public static void CopyAndPasteSize(PowerPoint.ShapeRange? shapeRange, CopyAndPasteCmd? cmd, Office.MsoScaleFrom scaleFromFlag, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (cmd == null) {
                return;
            }

            switch (cmd!) {
            case CopyAndPasteCmd.Copy:
                if (shapeRange.Count == 1) {
                    _copiedSizeWPt = shapeRange.Width;
                    _copiedSizeHPt = shapeRange.Height;
                    uiInvalidator?.Invoke();
                }
                break;
            case CopyAndPasteCmd.Paste:
                if (IsValidCopiedSizeValue()) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    foreach (var shape in shapeRange.OfType<PowerPoint.Shape>().ToArray()) {
                        shape.ScaleSizeTo(_copiedSizeWPt, _copiedSizeHPt, scaleFromFlag);
                    }
                    uiInvalidator?.Invoke();
                }
                break;
            case CopyAndPasteCmd.Reset:
                _copiedSizeWPt = InvalidCopiedValue;
                _copiedSizeHPt = InvalidCopiedValue;
                uiInvalidator?.Invoke();
                break;
            }
        }


        public static void CopyAndPastePosition(PowerPoint.ShapeRange? shapeRange, CopyAndPasteCmd? cmd, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (cmd == null) {
                return;
            }

            switch (cmd!) {
            case CopyAndPasteCmd.Copy:
                if (shapeRange.Count == 1) {
                    _copiedPositionXPt = shapeRange.Left;
                    _copiedPositionYPt = shapeRange.Top;
                    uiInvalidator?.Invoke();
                }
                break;
            case CopyAndPasteCmd.Paste:
                if (IsValidCopiedPositionValue()) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    shapeRange.Left = _copiedPositionXPt;
                    shapeRange.Top = _copiedPositionYPt;
                    uiInvalidator?.Invoke();
                }
                break;
            case CopyAndPasteCmd.Reset:
                _copiedPositionXPt = InvalidCopiedValue;
                _copiedPositionYPt = InvalidCopiedValue;
                uiInvalidator?.Invoke();
                break;
            }
        }

        public static bool IsSizeResettable(PowerPoint.ShapeRange? shapeRange) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return false;
            }

            var shapes = shapeRange.OfType<PowerPoint.Shape>().Where(s => s.Type is Office.MsoShapeType.msoPicture or Office.MsoShapeType.msoMedia).ToArray();
            return shapes.Length > 0;
        }

        public static void ResetMediaSize(PowerPoint.ShapeRange? shapeRange, Office.MsoScaleFrom? scaleFromFlag, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }

            var shapes = shapeRange.OfType<PowerPoint.Shape>().Where(s => s.Type is Office.MsoShapeType.msoPicture or Office.MsoShapeType.msoMedia).ToArray();
            if (shapes.Length == 0) {
                return;
            }

            const Office.MsoTriState relativeToOriginalSize = Office.MsoTriState.msoTrue;
            scaleFromFlag ??= Office.MsoScaleFrom.msoScaleFromTopLeft;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in shapes) {
                var isSound = shape.Type == Office.MsoShapeType.msoMedia && shape.MediaType == PowerPoint.PpMediaType.ppMediaTypeSound;
                var factor = !isSound ? 1F : 0.25F;
                shape.ScaleWidth(factor, relativeToOriginalSize, scaleFromFlag.Value);
                shape.ScaleHeight(factor, relativeToOriginalSize, scaleFromFlag.Value);
                shape.Rotation = 0;
                if (shape.HorizontalFlip == Office.MsoTriState.msoTrue) {
                    shape.Flip(Office.MsoFlipCmd.msoFlipHorizontal);
                }
                if (shape.VerticalFlip == Office.MsoTriState.msoTrue) {
                    shape.Flip(Office.MsoFlipCmd.msoFlipVertical);
                }
            }
            uiInvalidator?.Invoke();
        }

    }

}
