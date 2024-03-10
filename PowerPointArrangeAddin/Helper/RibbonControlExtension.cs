using System;
using Microsoft.Office.Core;

#pragma warning disable CS0618

namespace PowerPointArrangeAddin.Helper {

    public static class RibbonControlExtension {

        public static string TagOrId(this IRibbonControl ribbonControl) {
            // Microsoft.Office.Core.IRibbonControl;
            return !string.IsNullOrWhiteSpace(ribbonControl.Tag)
                ? ribbonControl.Tag
                : ribbonControl.Id;
        }

        private const string Separator = "·";

        public static string Id(this IRibbonControl ribbonControl) {
            var id = ribbonControl.Id;
            var parts = id.Split(new[] { Separator }, 2, StringSplitOptions.RemoveEmptyEntries);
            return parts[0];
        }

        public static string Group(this IRibbonControl ribbonControl) {
            var id = ribbonControl.Id;
            var parts = id.Split(new[] { Separator }, 2, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length < 2 ? "" : parts[1];
        }

        public static string CombineParentId(this IRibbonExtensibility _, string a, string b) {
            return $"{a}{Separator}{b}";
        }

        public static void InvalidateControl(this IRibbonUI ribbonUi, string controlId, string parentId) {
            ribbonUi.InvalidateControl($"{controlId}{Separator}{parentId}");
        }

        public static void InvalidateControl(this IRibbonUI ribbonUi, string controlId, (string, string) parentIds) {
            ribbonUi.InvalidateControl($"{controlId}{Separator}{parentIds.Item1}{Separator}{parentIds.Item2}");
        }

        public static void InvalidateControls(this IRibbonUI ribbonUi, string controlId, params object[] parentIds) {
            foreach (var parentId in parentIds) {
                switch (parentId) {
                case string id:
                    ribbonUi.InvalidateControl(controlId, id);
                    break;
                case (string id1, string id2):
                    ribbonUi.InvalidateControl(controlId, (id1, id2));
                    break;
                case string[] idArray:
                    foreach (var singleId in idArray) {
                        ribbonUi.InvalidateControl(controlId, singleId);
                    }
                    break;
                case (string, string)[] idsArray:
                    foreach (var (singleId1, singleId2) in idsArray) {
                        ribbonUi.InvalidateControl(controlId, (singleId1, singleId2));
                    }
                    break;
                }
            }
        }

    }

}
