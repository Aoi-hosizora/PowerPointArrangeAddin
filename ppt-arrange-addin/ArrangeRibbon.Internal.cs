using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using Office = Microsoft.Office.Core;

// #nullable enable

namespace ppt_arrange_addin {

    using RES = Properties.Resources;
    using ARES = ArrangeRibbonResources;

    [ComVisible(true)]
    public partial class ArrangeRibbon : Office.IRibbonExtensibility {

        public ArrangeRibbon() {
            InitializeAvailabilityRules();
        }

        public string GetCustomUI(string ribbonId) {
            var xml = GetResourceText("ppt_arrange_addin.ArrangeRibbon.xml");
            xml = ApplyTemplateForXml(xml);
            xml = ApplyMsoKeytipForXml(xml);
            return xml;
        }

        #region Helper Methods For GetCustomUI

        private static string GetResourceText(string resourceName) {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            foreach (var name in resourceNames) {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0) {
                    var stream = asm.GetManifestResourceStream(name);
                    if (stream != null) {
                        using var resourceReader = new StreamReader(stream);
                        return resourceReader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        private static string ApplyTemplateForXml(string xmlText) {
            var document = new XmlDocument();
            document.LoadXml(xmlText);

            // find templates node from document
            var templatesNodes = document.GetElementsByTagName("__templates");
            if (templatesNodes.Count == 0) {
                return document.OuterXml;
            }
            var templatesNode = templatesNodes[0];

            // extract templates to dictionary
            var templateDictionary = new Dictionary<string, Dictionary<string, string>>();
            foreach (var templateNode in templatesNode.ChildNodes.OfType<XmlNode>()) {
                var nodeAttributes = templateNode.Attributes;
                if (nodeAttributes == null) {
                    continue;
                }
                var name = nodeAttributes["name"]?.Value;
                if (string.IsNullOrWhiteSpace(name)) {
                    continue;
                }
                var attributes = new Dictionary<string, string>();
                foreach (var attribute in nodeAttributes.OfType<XmlAttribute>()) {
                    if (attribute.Name != "name") {
                        attributes[attribute.Name] = attribute.Value;
                    }
                }
                templateDictionary[name] = attributes;
            }
            templatesNode.ParentNode?.RemoveChild(templatesNode); // templates node must be removed

            // find nodes that need to be applied template
            var nodesToBeApplied = document.SelectNodes("//*[@__template]");
            if (nodesToBeApplied == null) {
                return document.OuterXml;
            }

            // apply template to each xml node
            foreach (var node in nodesToBeApplied.OfType<XmlNode>()) {
                var nodeAttributes = node.Attributes;
                var templateAttribute = nodeAttributes?["__template"];
                if (templateAttribute == null) {
                    continue;
                }

                nodeAttributes.RemoveNamedItem("__template"); // template attribute must be removed
                var templateNames = templateAttribute.Value?.Split(',');
                if (templateNames?.Length is null or 0) {
                    continue;
                }
                foreach (var templateName in templateNames) {
                    if (!templateDictionary.TryGetValue(templateName.Trim(), out var templateAttributes)) {
                        continue;
                    }
                    foreach (var attribute in templateAttributes) {
                        if (nodeAttributes[attribute.Key] != null) {
                            continue;
                        }
                        var newAttribute = document.CreateAttribute(attribute.Key);
                        newAttribute.Value = attribute.Value;
                        nodeAttributes.Append(newAttribute);
                    }
                }
            }

            // returned the applied xml string
            return document.OuterXml;
        }

        private class FakeRibbonControl : Office.IRibbonControl {
            public string Id { get; init; }
            public object Context => null;
            public string Tag => "";
        }

        private string ApplyMsoKeytipForXml(string xmlText) {
            var document = new XmlDocument();
            document.LoadXml(xmlText);

            // find nodes that are builtin controls
            var nodesToBeApplied = document.SelectNodes("//*[@idMso]");
            if (nodesToBeApplied == null) {
                return document.OuterXml;
            }

            string FindGroupId(XmlNode node) {
                var curr = node;
                while (curr != null) {
                    if (curr.Name == "group") {
                        return curr.Attributes?["id"]?.Value ?? "";
                    }
                    curr = curr.ParentNode;
                }
                return "";
            }

            // apply keytip to each xml node
            foreach (var node in nodesToBeApplied.OfType<XmlNode>()) {
                var nodeAttributes = node.Attributes;
                var idMsoValue = nodeAttributes?["idMso"]?.Value;
                var groupName = FindGroupId(node);
                if (string.IsNullOrWhiteSpace(idMsoValue) || string.IsNullOrWhiteSpace(groupName)) {
                    continue;
                }

                // check getKeytip and keytip attribute
                var getKeytip = nodeAttributes["getKeytip"]?.Value;
                var keytip = nodeAttributes["keytip"]?.Value;
                if (getKeytip != nameof(GetKeytip) || !string.IsNullOrWhiteSpace(keytip)) {
                    continue;
                }
                nodeAttributes.RemoveNamedItem("getKeytip");

                // query keytip of specific isMso control
                var keytipValue = GetKeytip(new FakeRibbonControl { Id = $"{groupName}.{idMsoValue}" });
                if (string.IsNullOrWhiteSpace(keytipValue)) {
                    continue;
                }

                // append new keytip attribute to node
                var keytipAttribute = document.CreateAttribute("keytip");
                keytipAttribute.Value = keytipValue;
                nodeAttributes.Append(keytipAttribute);
            }

            // returned the applied xml string
            return document.OuterXml;
        }

        #endregion

        #region Ribbon Elements ID

        // ReSharper disable InconsistentNaming
        // grpWordArt
        private const string grpWordArt = "grpWordArt";

        // grpArrange
        private const string grpArrange = "grpArrange";
        private const string btnAlignLeft = "btnAlignLeft";
        private const string btnAlignCenter = "btnAlignCenter";
        private const string btnAlignRight = "btnAlignRight";
        private const string btnAlignTop = "btnAlignTop";
        private const string btnAlignMiddle = "btnAlignMiddle";
        private const string btnAlignBottom = "btnAlignBottom";
        private const string btnDistributeHorizontal = "btnDistributeHorizontal";
        private const string btnDistributeVertical = "btnDistributeVertical";
        private const string btnScaleSameWidth = "btnScaleSameWidth";
        private const string btnScaleSameHeight = "btnScaleSameHeight";
        private const string btnScaleSameSize = "btnScaleSameSize";
        private const string btnScaleAnchor = "btnScaleAnchor";
        private const string btnExtendSameLeft = "btnExtendSameLeft";
        private const string btnExtendSameRight = "btnExtendSameRight";
        private const string btnExtendSameTop = "btnExtendSameTop";
        private const string btnExtendSameBottom = "btnExtendSameBottom";
        private const string btnSnapLeft = "btnSnapLeft";
        private const string btnSnapRight = "btnSnapRight";
        private const string btnSnapTop = "btnSnapTop";
        private const string btnSnapBottom = "btnSnapBottom";
        private const string btnMoveForward = "btnMoveForward";
        private const string btnMoveFront = "btnMoveFront";
        private const string btnMoveBackward = "btnMoveBackward";
        private const string btnMoveBack = "btnMoveBack";
        private const string btnRotateRight90 = "btnRotateRight90";
        private const string btnRotateLeft90 = "btnRotateLeft90";
        private const string btnFlipVertical = "btnFlipVertical";
        private const string btnFlipHorizontal = "btnFlipHorizontal";
        private const string btnGroup = "btnGroup";
        private const string btnUngroup = "btnUngroup";
        private const string mnuArrangement = "mnuArrangement";

        private const string btnAddInSetting = "btnAddInSetting";

        // grpTextbox
        private const string grpTextbox = "grpTextbox";
        private const string btnAutofitOff = "btnAutofitOff";
        private const string btnAutoShrinkText = "btnAutoShrinkText";
        private const string btnAutoResizeShape = "btnAutoResizeShape";
        private const string btnWrapText = "btnWrapText";
        private const string edtMarginLeft = "edtMarginLeft";
        private const string edtMarginRight = "edtMarginRight";
        private const string edtMarginTop = "edtMarginTop";
        private const string edtMarginBottom = "edtMarginBottom";
        private const string btnResetHorizontalMargin = "btnResetHorizontalMargin";

        private const string btnResetVerticalMargin = "btnResetVerticalMargin";

        // grpShapeSizeAndPosition
        private const string grpShapeSizeAndPosition = "grpShapeSizeAndPosition";
        private const string mnuShapeArrangement = "mnuShapeArrangement";
        private const string btnLockShapeAspectRatio = "btnLockShapeAspectRatio";
        private const string btnShapeScaleAnchor = "btnShapeScaleAnchor";
        private const string btnCopyShapeSize = "btnCopyShapeSize";
        private const string btnPasteShapeSize = "btnPasteShapeSize";
        private const string edtShapePositionX = "edtShapePositionX";
        private const string edtShapePositionY = "edtShapePositionY";
        private const string btnCopyShapePosition = "btnCopyShapePosition";

        private const string btnPasteShapePosition = "btnPasteShapePosition";

        // grpReplacePicture
        private const string grpReplacePicture = "grpReplacePicture";
        private const string btnReplaceWithClipboard = "btnReplaceWithClipboard";
        private const string btnReplaceWithFile = "btnReplaceWithFile";
        private const string chkReserveOriginalSize = "chkReserveOriginalSize";

        private const string chkReplaceToMiddle = "chkReplaceToMiddle";

        // grpPictureSizeAndPosition
        private const string grpPictureSizeAndPosition = "grpPictureSizeAndPosition";
        private const string mnuPictureArrangement = "mnuPictureArrangement";
        private const string btnResetPictureSize = "btnResetPictureSize";
        private const string btnLockPictureAspectRatio = "btnLockPictureAspectRatio";
        private const string btnPictureScaleAnchor = "btnPictureScaleAnchor";
        private const string btnCopyPictureSize = "btnCopyPictureSize";
        private const string btnPastePictureSize = "btnPastePictureSize";
        private const string edtPicturePositionX = "edtPicturePositionX";
        private const string edtPicturePositionY = "edtPicturePositionY";
        private const string btnCopyPicturePosition = "btnCopyPicturePosition";

        private const string btnPastePicturePosition = "btnPastePicturePosition";

        // mnuArrangement
        private const string mnuArrangement_sepAlignmentAndResizing = "mnuArrangement_sepAlignmentAndResizing";
        private const string mnuArrangement_mnuAlignment = "mnuArrangement_mnuAlignment";
        private const string mnuArrangement_mnuResizing = "mnuArrangement_mnuResizing";
        private const string mnuArrangement_mnuSnapping = "mnuArrangement_mnuSnapping";
        private const string mnuArrangement_mnuRotation = "mnuArrangement_mnuRotation";
        private const string mnuArrangement_sepLayerOrderAndGrouping = "mnuArrangement_sepLayerOrderAndGrouping";
        private const string mnuArrangement_mnuLayerOrder = "mnuArrangement_mnuLayerOrder";
        private const string mnuArrangement_mnuGrouping = "mnuArrangement_mnuGrouping";
        private const string mnuArrangement_sepObjectsInSlide = "mnuArrangement_sepObjectsInSlide";

        private const string mnuArrangement_sepAddInSetting = "mnuArrangement_sepAddInSetting";
        // ReSharper restore InconsistentNaming

        #endregion

        #region Element Ui Callbacks

        private class ElementUi {
            public string Label { get; init; }
            public System.Drawing.Image Image { get; init; }
            public string Keytip { get; init; }
        }

        private readonly Dictionary<string, Func<ElementUi>> _elementLabels = new() {
            // grpWordArt
            { grpWordArt, () => new ElementUi { Label = ARES.grpWordArt, Image = RES.TextEffectsMenu } },
            { "grpWordArt.TextStylesGallery", () => new ElementUi { Keytip = "AQ" } },
            { "grpWordArt.TextFillColorPicker", () => new ElementUi { Keytip = "AF" } },
            { "grpWordArt.TextOutlineColorPicker", () => new ElementUi { Keytip = "AU" } },
            { "grpWordArt.TextEffectsMenu", () => new ElementUi { Keytip = "AE" } },
            { "grpWordArt.WordArtFormatDialog", () => new ElementUi { Keytip = "AG" } },
            // grpArrange
            { grpArrange, () => new ElementUi { Label = ARES.grpArrange, Image = RES.ObjectArrangement } },
            { btnAlignLeft, () => new ElementUi { Label = ARES.btnAlignLeft, Image = RES.ObjectsAlignLeft, Keytip = "DL" } },
            { btnAlignCenter, () => new ElementUi { Label = ARES.btnAlignCenter, Image = RES.ObjectsAlignCenterHorizontal, Keytip = "DC" } },
            { btnAlignRight, () => new ElementUi { Label = ARES.btnAlignRight, Image = RES.ObjectsAlignRight, Keytip = "DR" } },
            { btnAlignTop, () => new ElementUi { Label = ARES.btnAlignTop, Image = RES.ObjectsAlignTop, Keytip = "DT" } },
            { btnAlignMiddle, () => new ElementUi { Label = ARES.btnAlignMiddle, Image = RES.ObjectsAlignMiddleVertical, Keytip = "DM" } },
            { btnAlignBottom, () => new ElementUi { Label = ARES.btnAlignBottom, Image = RES.ObjectsAlignBottom, Keytip = "DB" } },
            { "grpArrange.GridSettings", () => new ElementUi { Keytip = "DG" } },
            { btnDistributeHorizontal, () => new ElementUi { Label = ARES.btnDistributeHorizontal, Image = RES.AlignDistributeHorizontally, Keytip = "DH" } },
            { btnDistributeVertical, () => new ElementUi { Label = ARES.btnDistributeVertical, Image = RES.AlignDistributeVertically, Keytip = "DV" } },
            { btnScaleSameWidth, () => new ElementUi { Label = ARES.btnScaleSameWidth, Image = RES.ScaleSameWidth, Keytip = "PW" } },
            { btnScaleSameHeight, () => new ElementUi { Label = ARES.btnScaleSameHeight, Image = RES.ScaleSameHeight, Keytip = "PH" } },
            { btnScaleSameSize, () => new ElementUi { Label = ARES.btnScaleSameSize, Image = RES.ScaleSameSize, Keytip = "PS" } },
            { btnScaleAnchor, () => new ElementUi { Label = ARES.btnScaleAnchor_Middle, Image = RES.ScaleFromMiddle, Keytip = "PA" } },
            { btnExtendSameLeft, () => new ElementUi { Label = ARES.btnExtendSameLeft, Image = RES.ExtendSameLeft, Keytip = "PL" } },
            { btnExtendSameRight, () => new ElementUi { Label = ARES.btnExtendSameRight, Image = RES.ExtendSameRight, Keytip = "PR" } },
            { btnExtendSameTop, () => new ElementUi { Label = ARES.btnExtendSameTop, Image = RES.ExtendSameTop, Keytip = "PT" } },
            { btnExtendSameBottom, () => new ElementUi { Label = ARES.btnExtendSameBottom, Image = RES.ExtendSameBottom, Keytip = "PB" } },
            { btnSnapLeft, () => new ElementUi { Label = ARES.btnSnapLeft, Image = RES.SnapLeftToRight, Keytip = "PE" } },
            { btnSnapRight, () => new ElementUi { Label = ARES.btnSnapRight, Image = RES.SnapRightToLeft, Keytip = "PI" } },
            { btnSnapTop, () => new ElementUi { Label = ARES.btnSnapTop, Image = RES.SnapTopToBottom, Keytip = "PO" } },
            { btnSnapBottom, () => new ElementUi { Label = ARES.btnSnapBottom, Image = RES.SnapBottomToTop, Keytip = "PM" } },
            { btnMoveForward, () => new ElementUi { Label = ARES.btnMoveForward, Image = RES.ObjectBringForward, Keytip = "HF" } },
            { btnMoveFront, () => new ElementUi { Label = ARES.btnMoveFront, Image = RES.ObjectBringToFront, Keytip = "HO" } },
            { btnMoveBackward, () => new ElementUi { Label = ARES.btnMoveBackward, Image = RES.ObjectSendBackward, Keytip = "HB" } },
            { btnMoveBack, () => new ElementUi { Label = ARES.btnMoveBack, Image = RES.ObjectSendToBack, Keytip = "HK" } },
            { btnRotateRight90, () => new ElementUi { Label = ARES.btnRotateRight90, Image = RES.ObjectRotateRight90, Keytip = "HR" } },
            { btnRotateLeft90, () => new ElementUi { Label = ARES.btnRotateLeft90, Image = RES.ObjectRotateLeft90, Keytip = "HL" } },
            { btnFlipVertical, () => new ElementUi { Label = ARES.btnFlipVertical, Image = RES.ObjectFlipVertical, Keytip = "HV" } },
            { btnFlipHorizontal, () => new ElementUi { Label = ARES.btnFlipHorizontal, Image = RES.ObjectFlipHorizontal, Keytip = "HH" } },
            { btnGroup, () => new ElementUi { Label = ARES.btnGroup, Image = RES.ObjectsGroup, Keytip = "HG" } },
            { btnUngroup, () => new ElementUi { Label = ARES.btnUngroup, Image = RES.ObjectsUngroup, Keytip = "HU" } },
            { "grpArrange.ObjectSizeAndPositionDialog", () => new ElementUi { Keytip = "HS" } },
            { "grpArrange.SelectionPane", () => new ElementUi { Keytip = "HP" } },
            { mnuArrangement, () => new ElementUi { Label = ARES.mnuArrangement, Image = RES.ObjectArrangement_32, Keytip = "B" } },
            { btnAddInSetting, () => new ElementUi { Label = ARES.btnAddInSetting, Image = RES.AddInOptions, Keytip = "HT" } },
            // grpTextbox
            { grpTextbox, () => new ElementUi { Label = ARES.grpTextbox, Image = RES.TextboxSetting } },
            { btnAutofitOff, () => new ElementUi { Label = ARES.btnAutofitOff, Image = RES.TextboxAutofitOff, Keytip = "TF" } },
            { btnAutoShrinkText, () => new ElementUi { Label = ARES.btnAutoShrinkText, Image = RES.TextboxAutoShrinkText, Keytip = "TS" } },
            { btnAutoResizeShape, () => new ElementUi { Label = ARES.btnAutoResizeShape, Image = RES.TextboxAutoResizeShape, Keytip = "TR" } },
            { btnWrapText, () => new ElementUi { Label = ARES.btnWrapText, Image = RES.TextboxWrapText_32, Keytip = "TW" } },
            { edtMarginLeft, () => new ElementUi { Label = ARES.edtMarginLeft, Keytip = "ML" } },
            { edtMarginRight, () => new ElementUi { Label = ARES.edtMarginRight, Keytip = "MR" } },
            { edtMarginTop, () => new ElementUi { Label = ARES.edtMarginTop, Keytip = "MT" } },
            { edtMarginBottom, () => new ElementUi { Label = ARES.edtMarginBottom, Keytip = "MB" } },
            { btnResetHorizontalMargin, () => new ElementUi { Label = ARES.btnResetHorizontalMargin, Image = RES.TextboxResetMargin, Keytip = "MH" } },
            { btnResetVerticalMargin, () => new ElementUi { Label = ARES.btnResetVerticalMargin, Image = RES.TextboxResetMargin, Keytip = "MV" } },
            { "grpTextbox.WordArtFormatDialog", () => new ElementUi { Keytip = "TG" } },
            // grpShapeSizeAndPosition
            { grpShapeSizeAndPosition, () => new ElementUi { Label = ARES.grpShapeSizeAndPosition, Image = RES.SizeAndPosition } },
            { mnuShapeArrangement, () => new ElementUi { Label = ARES.mnuShapeArrangement, Image = RES.ObjectArrangement_32, Keytip = "B" } },
            { btnLockShapeAspectRatio, () => new ElementUi { Label = ARES.btnLockShapeAspectRatio, Image = RES.ObjectLockAspectRatio, Keytip = "L" } },
            { btnShapeScaleAnchor, () => new ElementUi { Label = ARES.btnScaleAnchor_Middle, Image = RES.ScaleFromMiddle, Keytip = "PA" } },
            { btnCopyShapeSize, () => new ElementUi { Label = ARES.btnCopyShapeSize, Image = RES.Copy, Keytip = "SC" } },
            { btnPasteShapeSize, () => new ElementUi { Label = ARES.btnPasteShapeSize, Image = RES.Paste, Keytip = "SP" } },
            { edtShapePositionX, () => new ElementUi { Label = ARES.edtShapePositionX, Keytip = "PX" } },
            { edtShapePositionY, () => new ElementUi { Label = ARES.edtShapePositionY, Keytip = "PY" } },
            { btnCopyShapePosition, () => new ElementUi { Label = ARES.btnCopyShapePosition, Image = RES.Copy, Keytip = "PC" } },
            { btnPasteShapePosition, () => new ElementUi { Label = ARES.btnPasteShapePosition, Image = RES.Paste, Keytip = "PP" } },
            { "grpShapeSizeAndPosition.ObjectSizeAndPositionDialog", () => new ElementUi { Keytip = "SN" } },
            // grpReplacePicture
            { grpReplacePicture, () => new ElementUi { Label = ARES.grpReplacePicture, Image = RES.PictureChangeFromClipboard } },
            { btnReplaceWithClipboard, () => new ElementUi { Label = ARES.btnReplaceWithClipboard, Image = RES.PictureChangeFromClipboard_32, Keytip = "TC" } },
            { btnReplaceWithFile, () => new ElementUi { Label = ARES.btnReplaceWithFile, Image = RES.PictureChange, Keytip = "TF" } },
            { chkReserveOriginalSize, () => new ElementUi { Label = ARES.chkReserveOriginalSize, Keytip = "TR" } },
            { chkReplaceToMiddle, () => new ElementUi { Label = ARES.chkReplaceToMiddle, Keytip = "TM" } },
            // grpPictureSizeAndPosition
            { grpPictureSizeAndPosition, () => new ElementUi { Label = ARES.grpPictureSizeAndPosition, Image = RES.SizeAndPosition } },
            { mnuPictureArrangement, () => new ElementUi { Label = ARES.mnuPictureArrangement, Image = RES.ObjectArrangement_32, Keytip = "B" } },
            { btnResetPictureSize, () => new ElementUi { Label = ARES.btnResetPictureSize, Image = RES.PictureResetSize_32, Keytip = "SR" } },
            { btnLockPictureAspectRatio, () => new ElementUi { Label = ARES.btnLockPictureAspectRatio, Image = RES.ObjectLockAspectRatio, Keytip = "L" } },
            { btnPictureScaleAnchor, () => new ElementUi { Label = ARES.btnScaleAnchor_Middle, Image = RES.ScaleFromMiddle, Keytip = "PA" } },
            { btnCopyPictureSize, () => new ElementUi { Label = ARES.btnCopyPictureSize, Image = RES.Copy, Keytip = "SC" } },
            { btnPastePictureSize, () => new ElementUi { Label = ARES.btnPastePictureSize, Image = RES.Paste, Keytip = "SP" } },
            { edtPicturePositionX, () => new ElementUi { Label = ARES.edtPicturePositionX, Keytip = "PX" } },
            { edtPicturePositionY, () => new ElementUi { Label = ARES.edtPicturePositionY, Keytip = "PY" } },
            { btnCopyPicturePosition, () => new ElementUi { Label = ARES.btnCopyPicturePosition, Image = RES.Copy, Keytip = "PC" } },
            { btnPastePicturePosition, () => new ElementUi { Label = ARES.btnPastePicturePosition, Image = RES.Paste, Keytip = "PP" } },
            { "grpPictureSizeAndPosition.ObjectSizeAndPositionDialog", () => new ElementUi { Keytip = "SN" } },
            // mnuArrangement
            { mnuArrangement_sepAlignmentAndResizing, () => new ElementUi { Label = ARES.mnuArrangement_sepAlignmentAndResizing } },
            { mnuArrangement_mnuAlignment, () => new ElementUi { Label = ARES.mnuArrangement_mnuAlignment, Image = RES.ObjectArrangement } },
            { mnuArrangement_mnuResizing, () => new ElementUi { Label = ARES.mnuArrangement_mnuResizing, Image = RES.ScaleSameWidth } },
            { mnuArrangement_mnuSnapping, () => new ElementUi { Label = ARES.mnuArrangement_mnuSnapping, Image = RES.SnapLeftToRight } },
            { mnuArrangement_mnuRotation, () => new ElementUi { Label = ARES.mnuArrangement_mnuRotation, Image = RES.ObjectRotateRight90 } },
            { mnuArrangement_sepLayerOrderAndGrouping, () => new ElementUi { Label = ARES.mnuArrangement_sepLayerOrderAndGrouping } },
            { mnuArrangement_mnuLayerOrder, () => new ElementUi { Label = ARES.mnuArrangement_mnuLayerOrder, Image = RES.ObjectSendToBack } },
            { mnuArrangement_mnuGrouping, () => new ElementUi { Label = ARES.mnuArrangement_mnuGrouping, Image = RES.ObjectsGroup } },
            { mnuArrangement_sepObjectsInSlide, () => new ElementUi { Label = ARES.mnuArrangement_sepObjectsInSlide } },
            { mnuArrangement_sepAddInSetting, () => new ElementUi { Label = ARES.mnuArrangement_sepAddInSetting } }
        };

        public string GetLabel(Office.IRibbonControl ribbonControl) {
            _elementLabels.TryGetValue(ribbonControl.Id, out var eui);
            return eui?.Invoke().Label ?? "<Unknown>";
        }

        public System.Drawing.Image GetImage(Office.IRibbonControl ribbonControl) {
            _elementLabels.TryGetValue(ribbonControl.Id, out var eui);
            return eui?.Invoke().Image;
        }

        public string GetKeytip(Office.IRibbonControl ribbonControl) {
            _elementLabels.TryGetValue(ribbonControl.Id, out var eui);
            return eui?.Invoke().Keytip ?? "";
        }

        #endregion

    }

}
