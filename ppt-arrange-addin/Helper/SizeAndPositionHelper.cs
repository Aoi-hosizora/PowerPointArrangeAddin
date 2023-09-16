using System;
using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

#nullable enable

namespace ppt_arrange_addin.Helper {

    public static class SizeAndPositionHelper {

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
                _ => -1
            };

            var text = "";
            if (valueInPt >= 0) {
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
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    _copiedSizeWPt = shapeRange.Width;
                    _copiedSizeHPt = shapeRange.Height;
                    uiInvalidator?.Invoke();
                }
                break;
            case CopyAndPasteCmd.Paste:
                if (IsValidCopiedSizeValue()) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    foreach (var shape in shapeRange.OfType<PowerPoint.Shape>().ToArray()) {
                        var oldLockSate = shape.LockAspectRatio;
                        shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                        var wRatio = _copiedSizeWPt / shape.Width;
                        shape.ScaleWidth(wRatio, Office.MsoTriState.msoFalse, scaleFromFlag);
                        var hRatio = _copiedSizeHPt / shape.Height;
                        shape.ScaleHeight(hRatio, Office.MsoTriState.msoFalse, scaleFromFlag);
                        shape.LockAspectRatio = oldLockSate;
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
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
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

        public static void ResetPictureSize(PowerPoint.ShapeRange? shapeRange, Office.MsoScaleFrom? scaleFromFlag, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }

            var pictures = shapeRange.OfType<PowerPoint.Shape>().Where((shape) => shape.Type == Office.MsoShapeType.msoPicture).ToArray();
            if (pictures.Length == 0) {
                return;
            }

            const Office.MsoTriState relativeToOriginalSize = Office.MsoTriState.msoTrue;
            scaleFromFlag ??= Office.MsoScaleFrom.msoScaleFromTopLeft;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in pictures) {
                if (shape.Type == Office.MsoShapeType.msoPicture) {
                    shape.ScaleWidth(1F, relativeToOriginalSize, scaleFromFlag.Value);
                    shape.ScaleHeight(1F, relativeToOriginalSize, scaleFromFlag.Value);
                }
            }
            uiInvalidator?.Invoke();
        }

    }

}
