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

        public static void SizeAndPositionDialog() {
            const string mso = "ObjectSizeAndPositionDialog";
            try {
                if (Globals.ThisAddIn.Application.CommandBars.GetEnabledMso(mso)) {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso(mso);
                }
            } catch (Exception) {
                // ignored
            }
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

        public enum SizeKind {
            Height,
            Width
        }

        public static void ChangeSizeOfString(PowerPoint.ShapeRange? shapeRange, SizeKind? sizeKind, Office.MsoScaleFrom? scaleFromFlag, string? input, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (sizeKind == null || input == null) {
                return;
            }
            var (valueInPt, ok) = UnitConverter.ParseStringToPtValue(input, canBeMinus: false);
            if (!ok) {
                uiInvalidator?.Invoke(); // reset input
                return;
            }

            scaleFromFlag ??= Office.MsoScaleFrom.msoScaleFromTopLeft;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            switch (sizeKind) {
            case SizeKind.Height:
                foreach (var shape in shapeRange.OfType<PowerPoint.Shape>().ToArray()) {
                    shape.ScaleHeightTo(valueInPt, scaleFromFlag.Value);
                }
                break;
            case SizeKind.Width:
                foreach (var shape in shapeRange.OfType<PowerPoint.Shape>().ToArray()) {
                    shape.ScaleWidthTo(valueInPt, scaleFromFlag.Value);
                }
                break;
            }
            uiInvalidator?.Invoke();
        }

        public static (string, bool) GetSizeOfString(PowerPoint.ShapeRange? shapeRange, SizeKind? sizeKind) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return ("", false);
            }
            if (sizeKind == null) {
                return ("", false);
            }

            var valueInPt = sizeKind! switch {
                SizeKind.Height => shapeRange.Height,
                SizeKind.Width => shapeRange.Width,
                _ => -1e9F // if shapes has different size, the value will be "-2.147484E+09"
            };

            var text = "";
            if (valueInPt >= -1e9F) {
                text = UnitConverter.FormatPtValueToString(valueInPt);
            }
            return (text, true);
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
            var (valueInPt, ok) = UnitConverter.ParseStringToPtValue(input, canBeMinus: true);
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

        private static float _copiedDistanceHPt = InvalidCopiedValue;
        private static float _copiedDistanceVPt = InvalidCopiedValue;

        public static bool IsSizeCopyable(PowerPoint.ShapeRange? shapeRange) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return false;
            }
            return shapeRange.Width >= -1e9F && shapeRange.Height >= -1e9F;
        }

        public static bool IsPositionCopyable(PowerPoint.ShapeRange? shapeRange) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return false;
            }
            return shapeRange.Left >= -1e9F && shapeRange.Top >= -1e9F;
        }

        public static bool IsValidCopiedSizeValue() {
            return !_copiedSizeWPt.Equals(InvalidCopiedValue) && !_copiedSizeHPt.Equals(InvalidCopiedValue);
        }

        public static bool IsValidCopiedPositionValue() {
            return !_copiedPositionXPt.Equals(InvalidCopiedValue) && !_copiedPositionYPt.Equals(InvalidCopiedValue);
        }

        public static bool IsValidCopiedDistanceHValue() {
            return !_copiedDistanceHPt.Equals(InvalidCopiedValue);
        }

        public static bool IsValidCopiedDistanceVValue() {
            return !_copiedDistanceVPt.Equals(InvalidCopiedValue);
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

        public enum DistanceType {
            RightLeft,
            LeftLeft,
            RightRight,
            LeftRight
        }

        public static void CopyAndPasteDistance(PowerPoint.ShapeRange? shapeRange, CopyAndPasteCmd? cmd, DistanceType? type, bool isHOrV, Action? uiInvalidator = null) {
            if (shapeRange == null) {
                return;
            }
            if (cmd == null) {
                return;
            }
            type ??= DistanceType.RightLeft;

            switch (cmd!) {
            case CopyAndPasteCmd.Copy: {
                var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
                if (shapes.Length != 2) {
                    return;
                }
                var (shape1, shape2) = (shapes[0], shapes[1]);
                var (left1, left2) = (shape1.Left, shape2.Left);
                var (width1, width2) = (shape1.Width, shape2.Width);
                var (top1, top2) = (shape1.Top, shape2.Top);
                var (height1, height2) = (shape1.Height, shape2.Height);
                bool seperated12, seperated21, intersected12, intersected21, contained12, contained21;
                if (isHOrV) {
                    seperated12 = left1 <= left2 && left1 + width1 <= left2;
                    seperated21 = left2 <= left1 && left2 + width2 <= left1;
                    intersected12 = left1 <= left2 && left1 + width1 >= left2 && left1 + width1 <= left2 + width2;
                    intersected21 = left2 <= left1 && left2 + width2 >= left1 && left2 + width2 <= left1 + width1;
                    contained12 = left1 <= left2 && left1 + width1 >= left2 + width2;
                    contained21 = left2 <= left1 && left2 + width2 >= left1 + width1;
                } else {
                    seperated12 = top1 <= top2 && top1 + height1 <= top2;
                    seperated21 = top2 <= top1 && top2 + height2 <= top1;
                    intersected12 = top1 <= top2 && top1 + height1 >= top2 && top1 + height1 <= top2 + height2;
                    intersected21 = top2 <= top1 && top2 + height2 >= top1 && top2 + height2 <= top1 + height1;
                    contained12 = top1 <= top2 && top1 + height1 >= top2 + height2;
                    contained21 = top2 <= top1 && top2 + height2 >= top1 + height1;
                }
                float distanceH = InvalidCopiedValue, distanceV = InvalidCopiedValue;
                switch (type!) {
                case DistanceType.LeftLeft:
                    if (!contained12 && !contained21) {
                        distanceH = Math.Abs(left1 - left2);
                        distanceV = Math.Abs(top1 - top2);
                    } else if (contained12) {
                        distanceH = left2 - left1;
                        distanceV = top2 - top1;
                    } else if (contained21) {
                        distanceH = left1 - left2;
                        distanceV = top1 - top2;
                    }
                    break;
                case DistanceType.RightRight:
                    if (!contained12 && !contained21) {
                        distanceH = Math.Abs((left1 + width1) - (left2 + width2));
                        distanceV = Math.Abs((top1 + height1) - (top2 + height2));
                    } else if (contained12) {
                        distanceH = (left2 + width2) - (left1 + width1);
                        distanceV = (top2 + height2) - (top1 + height1);
                    } else if (contained21) {
                        distanceH = (left1 + width1) - (left2 + width2);
                        distanceV = (top1 + height1) - (top2 + height2);
                    }
                    break;
                case DistanceType.RightLeft:
                    if (seperated12) {
                        distanceH = left2 - (left1 + width1);
                        distanceV = top2 - (top1 + height1);
                    } else if (seperated21) {
                        distanceH = left1 - (left2 + width2);
                        distanceV = top1 - (top2 + height2);
                    } else if (intersected12 || contained12) {
                        distanceH = left2 - (left1 + width1);
                        distanceV = top2 - (top1 + height1);
                    } else if (intersected21 || contained21) {
                        distanceH = left1 - (left2 + width2);
                        distanceV = top1 - (top2 + height2);
                    }
                    break;
                case DistanceType.LeftRight:
                    if (seperated12 || intersected12) {
                        distanceH = (left2 + width2) - left1;
                        distanceV = (top2 + height2) - top1;
                    } else if (seperated21 || intersected21) {
                        distanceH = (left1 + width1) - left2;
                        distanceV = (top1 + height1) - top2;
                    } else if (contained12) {
                        distanceH = (left2 + width2) - left1;
                        distanceV = (top2 + height2) - top1;
                    } else if (contained21) {
                        distanceH = (left1 + width1) - left2;
                        distanceV = (top1 + height1) - top2;
                    }
                    break;
                }
                if (isHOrV && !distanceH.Equals(InvalidCopiedValue)) {
                    _copiedDistanceHPt = distanceH;
                } else if (!isHOrV && !distanceV.Equals(InvalidCopiedValue)) {
                    _copiedDistanceVPt = distanceV;
                }
                uiInvalidator?.Invoke();
                break;
            }
            case CopyAndPasteCmd.Paste: {
                var shapes = shapeRange.OfType<PowerPoint.Shape>().ToArray();
                if (shapes.Length != 2) {
                    return;
                }
                var (shape1, shape2) = (shapes[0], shapes[1]);
                if (isHOrV && IsValidCopiedDistanceHValue()) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    // => move the second shape
                    shape2.Left = type! switch {
                        DistanceType.RightLeft => shape1.Left + shape1.Width + _copiedDistanceHPt,
                        DistanceType.LeftLeft => shape1.Left + _copiedDistanceHPt,
                        DistanceType.RightRight => shape1.Left + shape1.Width + _copiedDistanceHPt - shape2.Width,
                        DistanceType.LeftRight => shape1.Left + _copiedDistanceHPt - shape2.Width,
                        _ => shape2.Left
                    };
                    uiInvalidator?.Invoke();
                } else if (!isHOrV && IsValidCopiedDistanceVValue()) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    // => move the second shape
                    shape2.Top = type! switch {
                        DistanceType.RightLeft => shape1.Top + shape1.Height + _copiedDistanceVPt,
                        DistanceType.LeftLeft => shape1.Top + _copiedDistanceVPt,
                        DistanceType.RightRight => shape1.Top + shape1.Height + _copiedDistanceVPt - shape2.Height,
                        DistanceType.LeftRight => shape1.Top + _copiedDistanceVPt - shape2.Height,
                        _ => shape2.Top
                    };
                    uiInvalidator?.Invoke();
                }
                break;
            }
            case CopyAndPasteCmd.Reset:
                _copiedDistanceHPt = InvalidCopiedValue;
                _copiedDistanceVPt = InvalidCopiedValue;
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
