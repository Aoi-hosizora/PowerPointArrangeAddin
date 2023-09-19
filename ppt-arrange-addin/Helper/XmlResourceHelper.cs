using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using Office = Microsoft.Office.Core;

#nullable enable

namespace ppt_arrange_addin.Helper {

    public static class XmlResourceHelper {

        public static string? GetResourceText(string resourceName) {
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

        public static string ApplyTemplateForXml(string xmlText) {
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
                var name = nodeAttributes?["name"]?.Value;
                if (string.IsNullOrWhiteSpace(name)) {
                    continue;
                }
                var attributes = new Dictionary<string, string>();
                foreach (var attribute in nodeAttributes!.OfType<XmlAttribute>()) {
                    if (attribute.Name != "name") {
                        attributes[attribute.Name] = attribute.Value;
                    }
                }
                templateDictionary[name!] = attributes;
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

                nodeAttributes!.RemoveNamedItem("__template"); // template attribute must be removed
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

        public static string ApplyMsoKeytipForXml(string xmlText, Dictionary<string, Dictionary<string, string>> msoKeytips) {
            var document = new XmlDocument();
            document.LoadXml(xmlText);

            // find nodes that are builtin controls
            var nodesToBeApplied = document.SelectNodes("//*[@idMso]");
            if (nodesToBeApplied == null) {
                return document.OuterXml;
            }

            static string? FindGroupId(XmlNode node) {
                var curr = node;
                while (curr != null) {
                    if (curr.Name == "group") {
                        return curr.Attributes?["id"]?.Value ?? "";
                    }
                    curr = curr.ParentNode;
                }
                return null;
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
                var getKeytip = nodeAttributes!["getKeytip"]?.Value;
                if (string.IsNullOrWhiteSpace(getKeytip)) {
                    continue;
                }
                var keytip = nodeAttributes["keytip"]?.Value;
                if (!string.IsNullOrWhiteSpace(keytip)) {
                    continue;
                }

                // query keytip of specific isMso control
                var keytipValue = msoKeytips?[groupName!]?[idMsoValue!];
                if (string.IsNullOrWhiteSpace(keytipValue)) {
                    continue;
                }

                // append new keytip attribute to node
                var keytipAttribute = document.CreateAttribute("keytip");
                keytipAttribute.Value = keytipValue;
                nodeAttributes.Append(keytipAttribute);
                nodeAttributes.RemoveNamedItem("getKeytip"); // remove getKeytip attribute manually
            }

            // returned the applied xml string
            return document.OuterXml;
        }

    }

}
