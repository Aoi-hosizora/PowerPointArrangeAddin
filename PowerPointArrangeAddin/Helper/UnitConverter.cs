using System;
using System.Text.RegularExpressions;

#nullable enable

namespace PowerPointArrangeAddin.Helper {

    public static class UnitConverter {

        private static float CmToPt(float cm) => cm * 720F / 25.4F;

        private static float PtToCm(float pt) => pt * 25.4F / 720F;

        private static readonly Regex Re = new(@"^\s*(\d*\.?\d*)\s*(?:mm|cm)?\s*$", RegexOptions.IgnoreCase);

        public static (float, bool) ParseStringToPtValue(string text) {
            var matched = Re.Match(text);
            if (!matched.Success) {
                return (0, false);
            }

            var isMm = text.ToLower().Contains("mm");
            text = matched.Groups[1].Value;
            if (text.Length == 0) {
                text = "0";
            }
            if (!float.TryParse(text, out var valueInCm)) {
                return (0, false);
            }

            if (isMm) {
                valueInCm /= 10.0F;
            }
            var valueInPt = CmToPt(valueInCm);
            return (valueInPt, true);
        }

        public static string FormatPtValueToString(float pt) {
            var valueInCm = PtToCm(pt);
            return $"{Math.Round(valueInCm, 2)} cm";
        }

    }

}
