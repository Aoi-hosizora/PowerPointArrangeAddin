using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using Office = Microsoft.Office.Core;

#nullable enable

namespace PowerPointArrangeAddin.Helper {

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

        private const string Separator = "··";

        public static string NormalizeControlIdInGroup(string xmlText) {
            var document = new XmlDocument();
            try {
                document.LoadXml(xmlText);
            } catch (Exception) {
                return "";
            }

            static void Normalize(XmlNode node, string groupName) {
                if (node.Name == "group") {
                    return;
                }

                var attributes = node.Attributes;
                if (attributes != null) {
                    var idAttribute = attributes["id"];
                    var idValue = idAttribute?.Value;
                    if (idAttribute != null && !string.IsNullOrWhiteSpace(idValue)) {
                        idAttribute.Value = $"{idValue!}{Separator}{groupName}";
                    }
                }

                if (node.HasChildNodes) {
                    foreach (var childNode in node.ChildNodes.OfType<XmlNode>()) {
                        Normalize(childNode, groupName);
                    }
                }
            }

            var groupNodes = document.GetElementsByTagName("group");
            foreach (var groupNode in groupNodes.OfType<XmlNode>()) {
                var groupId = groupNode.Attributes?["id"]?.Value;
                if (string.IsNullOrWhiteSpace(groupId)) {
                    continue;
                }
                foreach (var childNode in groupNode.ChildNodes.OfType<XmlNode>()) {
                    Normalize(childNode, groupId!);
                }
            }

            return document.ToXmlText();
        }


        public static string NormalizeControlIdInMenu(string xmlText) {
            return xmlText;
        }

        private static string ToXmlText(this XmlNode node) {
            return node.OuterXml;
        }

        public static string ApplyAttributeTemplateForXml(string xmlText) {
            var document = new XmlDocument();
            try {
                document.LoadXml(xmlText);
            } catch (Exception) {
                return "";
            }

            // find templates node from document
            var templatesNodes = document.GetElementsByTagName("__templates");
            if (templatesNodes.Count == 0) {
                return document.ToXmlText();
            }
            var templatesNode = templatesNodes[0];

            // extract templates to dictionary (name to key-value pairs)
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
            if (nodesToBeApplied == null || nodesToBeApplied.Count == 0) {
                return document.ToXmlText();
            }

            // apply template to each xml node
            foreach (var node in nodesToBeApplied.OfType<XmlNode>()) {
                var nodeAttributes = node.Attributes;
                var templateAttributeValue = nodeAttributes?["__template"]?.Value;
                if (templateAttributeValue == null) {
                    continue;
                }
                nodeAttributes!.RemoveNamedItem("__template"); // template attribute must be removed
                var templateNames = templateAttributeValue.Split(','); // got template names
                if (templateNames.Length == 0) {
                    continue;
                }

                foreach (var templateName in templateNames) {
                    if (!templateDictionary.TryGetValue(templateName.Trim(), out var templateAttributes)) {
                        continue; // specific template is not found
                    }
                    foreach (var attribute in templateAttributes) {
                        if (nodeAttributes[attribute.Key] != null) {
                            continue; // only append attributes that is not contained previously
                        }
                        var newAttribute = document.CreateAttribute(attribute.Key);
                        newAttribute.Value = attribute.Value;
                        nodeAttributes.Append(newAttribute);
                    }
                }
            }

            // returned the applied xml string
            return document.ToXmlText();
        }

        public static string ApplySubtreeTemplateForXml(string xmlText) {
            var document = new XmlDocument();
            try {
                document.LoadXml(xmlText);
            } catch (Exception) {
                return "";
            }

            // find nodes that are subtree template
            var subtreeTemplateNodes = document.SelectNodes("//*[@__subtree_as_template]");
            if (subtreeTemplateNodes == null || subtreeTemplateNodes.Count == 0) {
                return document.ToXmlText();
            }

            // extract subtree templates to dictionary
            var templateDictionary = new Dictionary<string, XmlNodeList>();
            foreach (var templateNode in subtreeTemplateNodes.OfType<XmlNode>()) {
                var nodeAttributes = templateNode.Attributes;
                var templateName = nodeAttributes?["__subtree_as_template"]?.Value;
                nodeAttributes?.RemoveNamedItem("__subtree_as_template"); // template attribute must be removed 
                if (string.IsNullOrWhiteSpace(templateName)) {
                    continue;
                }
                templateDictionary[templateName!] = templateNode.ChildNodes;
            }

            // find nodes that need to be applied template
            var nodesToBeApplied = document.GetElementsByTagName("__apply_subtree_template").OfType<XmlNode>().ToArray();
            if (nodesToBeApplied.Length == 0) {
                return document.ToXmlText();
            }

            // apply subtree template to each xml node
            foreach (var node in nodesToBeApplied) {
                var parentNode = node.ParentNode;
                if (parentNode == null) {
                    continue; // almost unreachable
                }
                var nodeAttribute = node.Attributes;
                var (templateName, replaceField) = (nodeAttribute?["use_template"]?.Value, nodeAttribute?["replace_field"]?.Value);
                var (replaceFrom, replaceTo) = (nodeAttribute?["replace_from"]?.Value, nodeAttribute?["replace_to"]?.Value);
                parentNode.RemoveChild(node); // subtree template node must be removed
                if (string.IsNullOrWhiteSpace(templateName)) {
                    continue;
                }

                // get template node list, and construct children node dictionary
                if (!templateDictionary.TryGetValue(templateName!, out var templateNodeList)) {
                    continue; // specific template is not found
                }
                var parentChildrenNode = parentNode.ChildNodes.OfType<XmlNode>()
                    .ToDictionary(n => n.Name, n => n);

                static void ReplaceFieldValue(XmlNode newNode, string field, string from, string to) {
                    var attributes = newNode.Attributes;
                    var fieldValue = attributes?[field]?.Value;
                    if (!string.IsNullOrWhiteSpace(fieldValue)) {
                        attributes![field]!.Value = from switch {
                            "$" => attributes[field]!.Value + to,
                            "^" => to + attributes[field]!.Value,
                            _ => attributes[field]!.Value.Replace(from, to)
                        };
                    }
                    if (!newNode.HasChildNodes) {
                        return;
                    }
                    foreach (var childNode in newNode.ChildNodes.OfType<XmlNode>()) {
                        ReplaceFieldValue(childNode, field, from, to);
                    }
                }

                // enumerate template node list and apply to node 
                foreach (var templateNode in templateNodeList.OfType<XmlNode>()) {
                    var newNode = templateNode.CloneNode(true); // deep clone template node
                    if (!string.IsNullOrWhiteSpace(replaceField) && !string.IsNullOrWhiteSpace(replaceFrom) && replaceTo != null) {
                        ReplaceFieldValue(newNode, replaceField!, replaceFrom!, replaceTo);
                    }
                    if (!parentChildrenNode.ContainsKey(newNode.Name)) {
                        parentNode.AppendChild(newNode); // only append node if node name does not exist
                    }
                }
            }

            // returned the applied xml string
            return document.ToXmlText();
        }

        public static string ApplyMsoKeytipForXml(string xmlText, Dictionary<string, Dictionary<string, string>> msoKeytips) {
            var document = new XmlDocument();
            try {
                document.LoadXml(xmlText);
            } catch (Exception) {
                return "";
            }

            // find nodes that are builtin controls
            var nodesToBeApplied = document.SelectNodes("//*[@idMso]");
            if (nodesToBeApplied == null || nodesToBeApplied.Count == 0) {
                return document.ToXmlText();
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
                if (!msoKeytips.TryGetValue(groupName!, out var msoKeytipsMap)) {
                    continue;
                }
                if (!msoKeytipsMap.TryGetValue(idMsoValue!, out var keytipValue)) {
                    continue;
                }
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
            return document.ToXmlText();
        }

    }

}
