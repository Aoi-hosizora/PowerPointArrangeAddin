using System;
using System.Text.RegularExpressions;

#nullable enable

namespace PowerPointArrangeAddin.Helper {

    public static class UnitConverter {

        private static float CmToPt(float cm) => cm * 720F / 25.4F;

        private static float PtToCm(float pt) => pt * 25.4F / 720F;

        private static readonly Regex CmMmRe = new(@"^\s*[+-]?\s*(\d*\.?\d*)\s*(?:cm|mm)?\s*$", RegexOptions.IgnoreCase);

        private static readonly Regex DegreeRe = new(@"^\s*(\d*\.?\d*)\s*(?:°|度)?\s*$", RegexOptions.IgnoreCase);

        public static (float, bool) ParseStringToPtValue(string text, bool canBeMinus = false) {
            var matched = CmMmRe.Match(text);
            if (!matched.Success) {
                return (0, false);
            }

            var sign = text.Contains("-") ? -1 : 1;
            if (sign == -1 && !canBeMinus) {
                return (0, false);
            }

            var isMm = text.ToLower().Contains("mm");
            text = matched.Groups[1].Value;
            if (string.IsNullOrWhiteSpace(text)) {
                text = "0";
            }
            if (!float.TryParse(text, out var valueInCm)) {
                return (0, false);
            }

            if (isMm) {
                valueInCm /= 10.0F;
            }
            var valueInPt = CmToPt(valueInCm);
            return (sign * valueInPt, true);
        }

        public static string FormatPtValueToString(float pt) {
            var valueInCm = PtToCm(pt);
            return $"{Math.Round(valueInCm, 2)} cm";
        }

        public static (float, bool) ParseStringToDegValue(string text) {
            var matched = DegreeRe.Match(text);
            if (!matched.Success) {
                return (0, false);
            }

            text = matched.Groups[1].Value;
            if (string.IsNullOrWhiteSpace(text)) {
                text = "0";
            }
            if (!float.TryParse(text, out var valueInDeg)) {
                return (0, false);
            }
            return (valueInDeg, true);
        }

        public static string FormatDegValueToString(float deg) {
            return $"{Math.Round(deg, 1)}°";
        }

    }

}
