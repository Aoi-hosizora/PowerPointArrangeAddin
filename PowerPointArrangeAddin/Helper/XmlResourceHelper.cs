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

        public static string ApplyTemplateForXml(string xmlText) {
            xmlText = ApplyAttributeTemplateForXml(xmlText);
            xmlText = ApplySubtreeTemplateForXml(xmlText);
            return xmlText;
        }

        private static string ApplyAttributeTemplateForXml(string xmlText) {
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

        private static string ApplySubtreeTemplateForXml(string xmlText) {
            var document = Parse(xmlText);
            if (document == null) {
                return "";
            }
            var xmlns = new XmlNamespaceManager(document.NameTable);
            xmlns.AddNamespace("x", "http://schemas.microsoft.com/office/2006/01/customui");

            // find templates node, and nodes that are subtree template
            var templatesNodes = document.GetElementsByTagName("__templates").OfType<XmlNode>().ToArray();
            var subtreeTemplateNodes = document.SelectNodes("//*[@__as_subtree_template]")?.OfType<XmlNode>().ToArray() ?? new XmlNode[] { };

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

            // find nested templates
            var templateDependGraph = new Dictionary<string, HashSet<string>>();
            foreach (var kv in templateDictionary) {
                var (name, nodes) = (kv.Key, kv.Value);
                var dependNodes = nodes[0].ParentNode.SelectNodes(".//x:__use_template", xmlns)?.OfType<XmlNode>().ToArray();
                if (dependNodes == null || dependNodes.Length == 0) {
                    templateDependGraph[name] = new HashSet<string>();
                } else {
                    templateDependGraph[name] = dependNodes
                        .Select(n => n.Attributes?["name"]?.Value)
                        .Where(n => !string.IsNullOrWhiteSpace(n))
                        .Cast<string>().ToHashSet();
                }
            }
            while (templateDependGraph.Count != 0) {
                var freeTemplates = templateDependGraph.Where(kv => kv.Value.Count == 0).Select(kv => kv.Key).ToArray();
                if (freeTemplates.Length == 0) {
                    throw new Exception("There are dependency cycle in templates");
                }
                foreach (var freeTemplateName in freeTemplates) {
                    var dependFreeTemplateNames = templateDependGraph.Where(kv => kv.Value.Contains(freeTemplateName)).Select(kv => kv.Key).ToArray();
                    foreach (var templateName in dependFreeTemplateNames) {
                        var templateNodes = templateDictionary[templateName].OfType<XmlNode>().ToList();
                        var needToBePrepared = templateNodes[0].ParentNode.SelectNodes($".//*[@name=\"{freeTemplateName}\"]")?.OfType<XmlNode>();
                        if (needToBePrepared != null) {
                            foreach (var node in needToBePrepared) {
                                ApplySingleNode(document, node, templateDictionary);
                            }
                        }
                        templateDependGraph[templateName].Remove(freeTemplateName);
                    }
                    templateDependGraph.Remove(freeTemplateName);
                }
            }

            // find nodes that are need to be applied template
            // apply subtree template to each xml node
            var nodesToBeApplied = document.GetElementsByTagName("__use_template").OfType<XmlNode>().ToList();
            if (nodesToBeApplied.Count > 0) {
                foreach (var node in nodesToBeApplied) {
                    ApplySingleNode(document, node, templateDictionary);
                }
            }


            nodesToBeApplied = document.GetElementsByTagName("__use_reference").OfType<XmlNode>().ToList();
            if (nodesToBeApplied.Count > 0) {
                foreach (var node in nodesToBeApplied) {
                    ApplySingleNode(document, node, templateDictionary);
                }
            }

            static void ApplySingleNode(XmlDocument document, XmlNode node, IReadOnlyDictionary<string, XmlNodeList> templateDictionary) {
                // get template name and two rule lists
                var templateName = node.Attributes?["name"]?.Value;
                if (string.IsNullOrWhiteSpace(templateName)) {
                    return;
                }
                var replaceRules = new List<(string field, string from, string to, bool norec, bool re)>();
                var removeRules = new List<(string field, string match, bool norec, bool re)>();
                var ruleNodes = node.ChildNodes.OfType<XmlNode>().ToList();
                ruleNodes.Add(node); // also regard the use_template node as rule node
                foreach (var ruleNode in ruleNodes) {
                    var nodeName = ruleNode.Name;
                    var nodeAttributes = ruleNode.Attributes;
                    var fieldValue = nodeAttributes?["field"]?.Value;
                    var norecValue = nodeAttributes?["norec"]?.Value == "true";
                    if (nodeName == node.Name) {
                        // allow to define rule on template node directly 
                        var (replaceFieldValue, removeFieldValue) = (nodeAttributes?["replace_rule_field"]?.Value, nodeAttributes?["remove_rule_field"]?.Value);
                        if (!string.IsNullOrWhiteSpace(replaceFieldValue)) {
                            (nodeName, fieldValue) = ("__replace_rule", replaceFieldValue);
                        } else if (!string.IsNullOrWhiteSpace(removeFieldValue)) {
                            (nodeName, fieldValue) = ("__remove_rule", removeFieldValue);
                        } else {
                            continue;
                        }
                    }
                    switch (nodeName) {
                    case "__replace_rule":
                        var (fromValue, fromReValue, toValue) = (nodeAttributes?["from"]?.Value, nodeAttributes?["from_re"]?.Value, nodeAttributes?["to"]?.Value);
                        if (string.IsNullOrWhiteSpace(fieldValue) || (string.IsNullOrWhiteSpace(fromValue) && string.IsNullOrWhiteSpace(fromReValue)) || toValue == null) {
                            continue;
                        }
                        if (!string.IsNullOrWhiteSpace(fromReValue) is var useReForReplacing && useReForReplacing) {
                            fromValue = fromReValue!; // use from_re instead
                        }
                        replaceRules.Add((fieldValue!, fromValue!, toValue, norecValue, useReForReplacing));
                        break;
                    case "__remove_rule":
                        var (matchValue, matchReValue) = (nodeAttributes?["match"]?.Value, nodeAttributes?["match_re"]?.Value);
                        if (string.IsNullOrWhiteSpace(fieldValue) || (string.IsNullOrWhiteSpace(matchValue) && string.IsNullOrWhiteSpace(matchReValue))) {
                            continue;
                        }
                        if (!string.IsNullOrWhiteSpace(matchReValue) is var useReForRemoving && useReForRemoving) {
                            matchValue = matchReValue!; // use match_re instead
                        }
                        removeRules.Add((fieldValue!, matchValue!, norecValue, useReForRemoving));
                        break;
                    default:
                        continue;
                    }
                }

                // get template subtree (node list) or ref control (single node)
                List<XmlNode> templateNodeList;
                switch (node.Name) {
                case "__use_template":
                    if (!templateDictionary.TryGetValue(templateName!, out var nodeList)) {
                        return; // specific template is not found
                    }
                    templateNodeList = nodeList.OfType<XmlNode>().ToList();
                    break;
                case "__use_reference":
                    var matched = new Regex(@"^\$(\w+?)=(\w+?)(?::(\d+?))?$").Match(templateName!);
                    if (!matched.Success) {
                        return; // failed to extract referee
                    }
                    var (keyValue, valueValue, numValue) = (matched.Groups[1].Value, matched.Groups[2].Value, matched.Groups[3].Value);
                    if (string.IsNullOrWhiteSpace(keyValue) || string.IsNullOrWhiteSpace(valueValue)) {
                        return; // use empty key or value in template
                    }
                    if (!int.TryParse(string.IsNullOrWhiteSpace(numValue) ? "1" : numValue!, out var num)) {
                        return; // use invalid num string value
                    }
                    var foundNodes = document.SelectNodes($"//*[@{keyValue}=\"{valueValue}\"]")?.OfType<XmlNode>().ToArray();
                    if (foundNodes == null || foundNodes.Length == 0) {
                        return; // can not find template referee node
                    }
                    num = Math.Min(1, Math.Max(foundNodes.Length, num));
                    templateNodeList = new List<XmlNode> { foundNodes[num - 1] };
                    break;
                default:
                    return;
                }

                // two rule functions
                static void ReplaceFieldValue(XmlNode node, List<(string, string, string, bool, bool)> replaceRules, int layer = 1) {
                    var attributes = node.Attributes;
                    foreach (var (field, from, to, norec, re) in replaceRules) {
                        if (norec && layer > 1) {
                            continue;
                        }
                        var fieldValue = attributes?[field]?.Value ?? "";
                        if (!re) {
                            fieldValue = fieldValue.Replace(from, to);
                        } else {
                            try {
                                fieldValue = new Regex(from).Replace(fieldValue, to);
                            } catch (Exception) { }
                        }
                        attributes?.RemoveNamedItem(field);
                        if (!string.IsNullOrWhiteSpace(fieldValue)) {
                            attributes?.Append(node.CreateAttributeWithValue(field, fieldValue));
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

        public static string ApplyControlRandomId(string xmlText) {
            var document = Parse(xmlText);
            if (document == null) {
                return "";
            }

            var nodesToBeApplied = document.SelectNodes("//*[@id=\"*\"]");
            if (nodesToBeApplied == null || nodesToBeApplied.Count == 0) {
                return document.ToXmlText();
            }

            foreach (var node in nodesToBeApplied.OfType<XmlNode>()) {
                var idAttribute = node.Attributes?["id"];
                var idValue = idAttribute?.Value;
                if (idAttribute == null || idValue != "*") {
                    continue;
                }
                idAttribute.Value = "_" + Guid.NewGuid().ToString().Replace("-", "");
            }

            return document.ToXmlText();
        }

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
