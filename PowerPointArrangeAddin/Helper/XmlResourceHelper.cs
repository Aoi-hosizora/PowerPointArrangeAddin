using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
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

        private static XmlDocument? Parse(string xmlText) {
            var settings = new XmlReaderSettings { IgnoreComments = true };
            var document = new XmlDocument();
            try {
                var reader = new StringReader(xmlText);
                var xmlReader = XmlReader.Create(reader, settings);
                document.Load(xmlReader);
            } catch (Exception) {
                return null;
            }
            return document;
        }

        private static string ToXmlText(this XmlNode node) {
            return node.OuterXml;
        }

        private static XmlAttribute CreateAttributeWithValue(this XmlNode node, string key, string value) {
            if (node is not XmlDocument doc) {
                doc = node.OwnerDocument!;
            }
            var attribute = doc.CreateAttribute(key);
            attribute.Value = value;
            return attribute;
        }

        #region Normalize Related

        private const string Separator = "·";

        private static void NormalizeControlIdWithXmlNode(XmlNode node, string suffix, Func<XmlNode, bool>? toBreak) {
            if (toBreak != null && toBreak.Invoke(node)) {
                return;
            }

            var attributes = node.Attributes;
            if (attributes != null) {
                var idAttribute = attributes["id"];
                var idValue = idAttribute?.Value;
                if (idAttribute != null && !string.IsNullOrWhiteSpace(idValue)) {
                    idAttribute.Value = $"{idValue!}{Separator}{suffix}";
                }
            }

            if (node.HasChildNodes) {
                foreach (var childNode in node.ChildNodes.OfType<XmlNode>()) {
                    NormalizeControlIdWithXmlNode(childNode, suffix, toBreak);
                }
            }
        }

        public static string NormalizeControlIdInGroup(string xmlText) {
            if (xmlText.Contains(Separator)) {
                throw new Exception($"`{Separator}` can not be used in resource xml");
            }
            var document = Parse(xmlText);
            if (document == null) {
                return "";
            }

            var groupNodes = document.GetElementsByTagName("group");
            foreach (var groupNode in groupNodes.OfType<XmlNode>()) {
                var groupId = groupNode.Attributes?["id"]?.Value;
                if (string.IsNullOrWhiteSpace(groupId)) {
                    continue;
                }
                foreach (var childNode in groupNode.ChildNodes.OfType<XmlNode>()) {
                    NormalizeControlIdWithXmlNode(childNode, groupId!, node => node.Name == "group");
                }
            }
            return document.ToXmlText();
        }

        public static string NormalizeControlIdInMenu(string xmlText, string menuId) {
            if (xmlText.Contains(Separator)) {
                throw new Exception($"`{Separator}` can not be used in resource xml");
            }
            var document = Parse(xmlText);
            if (document == null) {
                return "";
            }

            var menuRootNode = document.GetElementsByTagName("menu").OfType<XmlNode>().FirstOrDefault();
            if (menuRootNode == null) {
                return document.ToXmlText();
            }

            foreach (var childNode in menuRootNode.ChildNodes.OfType<XmlNode>()) {
                NormalizeControlIdWithXmlNode(childNode, menuId, null);
            }
            return document.ToXmlText();
        }

        #endregion

        #region Template Related

        public static string ApplyAttributeTemplateForXml(string xmlText) {
            var document = Parse(xmlText);
            if (document == null) {
                return "";
            }

            // find templates node from document
            var templatesNodes = document.GetElementsByTagName("__templates").OfType<XmlNode>().ToArray();
            if (templatesNodes.Length == 0) {
                return document.ToXmlText();
            }

            // extract attribute templates to dictionary (template name to key-value pairs)
            var templateDictionary = new Dictionary<string, Dictionary<string, string>>();
            foreach (var templatesNode in templatesNodes) {
                var templateNodes = templatesNode.ChildNodes.OfType<XmlNode>().ToArray();
                foreach (var templateNode in templateNodes) {
                    if (templateNode.Name != "__attribute_template") {
                        continue;
                    }
                    var nodeAttributes = templateNode.Attributes;
                    var nameValue = nodeAttributes?["name"]?.Value;
                    if (string.IsNullOrWhiteSpace(nameValue)) {
                        continue;
                    }
                    var attributes = new Dictionary<string, string>();
                    foreach (var attribute in nodeAttributes!.OfType<XmlAttribute>()) {
                        if (attribute.Name != "name") {
                            attributes[attribute.Name] = attribute.Value;
                        }
                    }
                    templateDictionary[nameValue!] = attributes;
                    templateNode.ParentNode?.RemoveChild(templateNode); // template node must be removed
                }
                if (!templatesNode.HasChildNodes) {
                    templatesNode.ParentNode?.RemoveChild(templatesNode); // templates node must be removed    
                }
            }

            // find nodes that are need to be applied template
            var nodesToBeApplied = document.SelectNodes("//*[@__template]");
            if (nodesToBeApplied == null || nodesToBeApplied.Count == 0) {
                return document.ToXmlText();
            }

            // apply attribute template to each xml node
            foreach (var node in nodesToBeApplied.OfType<XmlNode>()) {
                var nodeAttributes = node.Attributes;
                var templateValue = nodeAttributes?["__template"]?.Value;
                if (templateValue == null) {
                    continue;
                }
                nodeAttributes!.RemoveNamedItem("__template"); // template attribute must be removed
                var templateNames = templateValue.Split(','); // get template name array
                if (templateNames.Length == 0) {
                    continue;
                }

                foreach (var templateName in templateNames) {
                    if (!templateDictionary.TryGetValue(templateName.Trim(), out var templateAttributes)) {
                        continue; // specific template is not found
                    }
                    foreach (var attribute in templateAttributes) {
                        if (nodeAttributes[attribute.Key] != null) {
                            continue; // only append attributes that are not defined
                        }
                        var newAttribute = document.CreateAttributeWithValue(attribute.Key, attribute.Value);
                        nodeAttributes.Append(newAttribute);
                    }
                }
            }

            // returned the applied xml string
            return document.ToXmlText();
        }

        public static string ApplySubtreeTemplateForXml(string xmlText) {
            var document = Parse(xmlText);
            if (document == null) {
                return "";
            }

            // find templates node, and nodes that are subtree template
            var templatesNodes = document.GetElementsByTagName("__templates").OfType<XmlNode>().ToArray();
            var subtreeTemplateNodes = document.SelectNodes("//*[@__as_subtree_template]")?.OfType<XmlNode>().ToArray() ?? new XmlNode[] { };
            if (templatesNodes.Length == 0 && subtreeTemplateNodes.Length == 0) {
                return document.ToXmlText();
            }

            // extract subtree templates to dictionary (template name to xml node list)
            var templateDictionary = new Dictionary<string, XmlNodeList>();
            foreach (var templatesNode in templatesNodes) {
                var templateNodes = templatesNode.ChildNodes.OfType<XmlNode>().ToArray();
                foreach (var templateNode in templateNodes) {
                    if (templateNode.Name != "__subtree_template") {
                        continue;
                    }
                    var nodeAttributes = templateNode.Attributes;
                    var nameValue = nodeAttributes?["name"]?.Value;
                    if (string.IsNullOrWhiteSpace(nameValue)) {
                        continue;
                    }
                    templateDictionary[nameValue!] = templateNode.ChildNodes;
                    templateNode.ParentNode?.RemoveChild(templateNode); // template node must be removed
                }
                if (!templatesNode.HasChildNodes) {
                    templatesNode.ParentNode?.RemoveChild(templatesNode); // templates node must be removed    
                }
            }
            foreach (var templateNode in subtreeTemplateNodes) {
                var nodeAttributes = templateNode.Attributes;
                var nameValue = nodeAttributes?["__as_subtree_template"]?.Value;
                nodeAttributes?.RemoveNamedItem("__as_subtree_template"); // template attribute must be removed 
                if (string.IsNullOrWhiteSpace(nameValue)) {
                    continue;
                }
                templateDictionary[nameValue!] = templateNode.ChildNodes;
            }

            // find nodes that are need to be applied template
            var nodesToBeApplied = document.GetElementsByTagName("__use_template").OfType<XmlNode>().ToArray();
            if (nodesToBeApplied.Length == 0) {
                return document.ToXmlText();
            }

            // apply subtree template to each xml node
            foreach (var node in nodesToBeApplied) {
                // get template name and two rule lists
                var templateName = node.Attributes?["name"]?.Value;
                if (string.IsNullOrWhiteSpace(templateName)) {
                    continue;
                }
                var replaceRules = new List<(string field, string from, string to, bool norec, bool re)>();
                var removeRules = new List<(string field, string match, bool norec, bool re)>();
                foreach (var ruleNode in node.ChildNodes.OfType<XmlNode>().ToArray()) {
                    var nodeAttributes = ruleNode.Attributes;
                    var fieldValue = nodeAttributes?["field"]?.Value;
                    var norecValue = nodeAttributes?["norec"]?.Value == "true";
                    switch (ruleNode.Name) {
                    case "__replace_rule":
                        var (fromValue, fromReValue, toValue) = (nodeAttributes?["from"]?.Value, nodeAttributes?["from_re"]?.Value, nodeAttributes?["to"]?.Value);
                        if (string.IsNullOrWhiteSpace(fieldValue) || string.IsNullOrWhiteSpace(toValue)) {
                            continue;
                        }
                        if (string.IsNullOrWhiteSpace(fromValue) && string.IsNullOrWhiteSpace(fromReValue)) {
                            continue;
                        }
                        if (!string.IsNullOrWhiteSpace(fromReValue) is var useReForReplacing && useReForReplacing) {
                            fromValue = fromReValue!;
                        }
                        replaceRules.Add((fieldValue!, fromValue!, toValue!, norecValue, useReForReplacing));
                        break;
                    case "__remove_rule":
                        var (matchValue, matchReValue) = (nodeAttributes?["match"]?.Value, nodeAttributes?["match_re"]?.Value);
                        if (string.IsNullOrWhiteSpace(fieldValue)) {
                            continue;
                        }
                        if (string.IsNullOrWhiteSpace(matchValue) && string.IsNullOrWhiteSpace(matchReValue)) {
                            continue;
                        }
                        if (!string.IsNullOrWhiteSpace(matchReValue) is var useReForRemoving && useReForRemoving) {
                            matchValue = matchReValue!;
                        }
                        removeRules.Add((fieldValue!, matchValue!, norecValue, useReForRemoving));
                        break;
                    }
                }

                // get template subtree (node list) or ref control (single node)
                List<XmlNode> templateNodeList;
                var matched = new Regex(@"^\$(\w+?)=(\w+?)(?::(\d+?))?$").Match(templateName!);
                if (!matched.Success) {
                    if (!templateDictionary.TryGetValue(templateName!, out var nodeList)) {
                        continue; // specific template is not found
                    }
                    templateNodeList = nodeList.OfType<XmlNode>().ToList();
                } else {
                    var (keyValue, valueValue, numValue) = (matched.Groups[1].Value, matched.Groups[2].Value, matched.Groups[3].Value);
                    if (string.IsNullOrWhiteSpace(keyValue) || string.IsNullOrWhiteSpace(valueValue)) {
                        continue; // use empty key or value in template
                    }
                    if (!int.TryParse(string.IsNullOrWhiteSpace(numValue) ? "1" : numValue!, out var num)) {
                        continue; // use invalid num string value
                    }
                    var foundNodes = document.SelectNodes($"//*[@{keyValue}=\"{valueValue}\"]")?.OfType<XmlNode>().ToArray();
                    if (foundNodes == null || foundNodes.Length == 0) {
                        continue; // can not find template referee node
                    }
                    num = Math.Min(1, Math.Max(foundNodes.Length, num));
                    templateNodeList = new List<XmlNode> { foundNodes[num - 1] };
                }

                // two rule functions
                static void ReplaceFieldValue(XmlNode node, List<(string, string, string, bool, bool)> replaceRules, int layer = 1) {
                    var attributes = node.Attributes;
                    foreach (var (field, from, to, norec, re) in replaceRules) {
                        if (norec && layer > 1) {
                            continue;
                        }
                        var fieldValue = attributes?[field]?.Value;
                        if (string.IsNullOrWhiteSpace(fieldValue)) { // insert attribute
                            attributes?.RemoveNamedItem(field);
                            attributes?.Append(node.CreateAttributeWithValue(field, to));
                        } else { // update attribute
                            if (!re) {
                                fieldValue = fieldValue!.Replace(from, to);
                            } else {
                                try {
                                    fieldValue = new Regex(from).Replace(fieldValue!, to);
                                } catch (Exception) { }
                            }
                            attributes![field]!.Value = fieldValue;
                        }
                    }
                    foreach (var childNode in node.ChildNodes.OfType<XmlNode>()) {
                        ReplaceFieldValue(childNode, replaceRules, layer + 1);
                    }
                }
                static XmlNode? RemoveSpecificNode(XmlNode node, List<(string, string, bool, bool)> removeRules, int layer = 1) {
                    var attributes = node.Attributes;
                    foreach (var (field, match, norec, re) in removeRules) {
                        if (norec && layer > 1) {
                            continue;
                        }
                        var fieldValue = field == "@" ? node.Name : attributes?[field]?.Value;
                        fieldValue ??= ""; // also allow for checking empty
                        var matched = false;
                        if (!re) {
                            matched = fieldValue.Contains(match);
                        } else {
                            try {
                                matched = new Regex(match).IsMatch(fieldValue);
                            } catch (Exception) { }
                        }
                        if (matched) {
                            return null; // delete this node
                        }
                    }
                    foreach (var childNode in node.ChildNodes.OfType<XmlNode>().ToArray()) { // need to copy here, because of `RemoveChild`
                        var result = RemoveSpecificNode(childNode, removeRules, layer + 1);
                        if (result == null) {
                            node.RemoveChild(childNode);
                        }
                    }
                    return node;
                }

                // ==> enumerate template node list and apply to node 
                foreach (var templateNode in templateNodeList.OfType<XmlNode>()) {
                    var clonedNode = templateNode.CloneNode(true); // deep clone template node
                    if (replaceRules.Count > 0) {
                        ReplaceFieldValue(clonedNode, replaceRules);
                    }
                    if (removeRules.Count > 0) {
                        clonedNode = RemoveSpecificNode(clonedNode, removeRules);
                    }
                    if (clonedNode != null) {
                        node.ParentNode?.InsertBefore(clonedNode, node);
                    }
                }
                node.ParentNode?.RemoveChild(node); // subtree template node must be removed
            }

            // returned the applied xml string
            return document.ToXmlText();
        }

        #endregion

        #region Misc Methods

        public static string ApplyMsoKeytipForXml(string xmlText, Dictionary<string, Dictionary<string, string>> msoKeytips) {
            var document = Parse(xmlText);
            if (document == null) {
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
                var keytipAttribute = document.CreateAttributeWithValue("keytip", keytipValue);
                nodeAttributes.Append(keytipAttribute);
                nodeAttributes.RemoveNamedItem("getKeytip"); // remove getKeytip attribute manually
            }

            // returned the applied xml string
            return document.ToXmlText();
        }

        #endregion

    }

}
