using System;
using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

#nullable enable

namespace PowerPointArrangeAddin.Helper {

    public static class RotationHelper {

        public static void ChangeAngleOfString(PowerPoint.ShapeRange? shapeRange, string? input, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (input == null) {
                return;
            }
            var (valueInDef, ok) = UnitConverter.ParseStringToDegValue(input);
            if (!ok) {
                uiInvalidator?.Invoke(); // reset input
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            shapeRange.Rotation = valueInDef;
            uiInvalidator?.Invoke();
        }

        public static (string, bool) GetAngleOfString(PowerPoint.ShapeRange? shapeRange) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return ("", false);
            }

            var valueInDeg = shapeRange.Rotation; // if shapes has different angle, the value will be "-2.147484E+09"

            var text = "";
            if (valueInDeg >= -1e9F) {
                text = UnitConverter.FormatDegValueToString(valueInDeg);
            }
            return (text, true);
        }

        private const float InvalidCopiedValue = -2147483648.0F; // for angle

        private static float _copiedAngleDeg = InvalidCopiedValue;

        public static bool IsAngleCopyable(PowerPoint.ShapeRange? shapeRange) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return false;
            }
            return shapeRange.Rotation >= -1e9F;
        }

        public static bool IsValidCopiedAngleValue() {
            return !_copiedAngleDeg.Equals(InvalidCopiedValue);
        }

        public enum CopyAndPasteCmd {
            Copy,
            Paste,
            Reset
        }

        public static void CopyAndPasteAngle(PowerPoint.ShapeRange? shapeRange, CopyAndPasteCmd? cmd, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            if (cmd == null) {
                return;
            }

            switch (cmd!) {
            case CopyAndPasteCmd.Copy:
                if (shapeRange.Count == 1) {
                    _copiedAngleDeg = shapeRange.Rotation;
                    uiInvalidator?.Invoke();
                }
                break;
            case CopyAndPasteCmd.Paste:
                if (IsValidCopiedAngleValue()) {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                    foreach (var shape in shapeRange.OfType<PowerPoint.Shape>().ToArray()) {
                        shape.Rotation = _copiedAngleDeg;
                    }
                    uiInvalidator?.Invoke();
                }
                break;
            case CopyAndPasteCmd.Reset:
                _copiedAngleDeg = InvalidCopiedValue;
                uiInvalidator?.Invoke();
                break;
            }
        }

        public static void ResetObjectAngle(PowerPoint.ShapeRange? shapeRange, Action? uiInvalidator = null) {
            if (shapeRange == null || shapeRange.Count <= 0) {
                return;
            }
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            shapeRange.Rotation = 0;
            uiInvalidator?.Invoke();
        }

    }

}
