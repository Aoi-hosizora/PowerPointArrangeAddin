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

        #region Normalize Related

        private const string Separator = "·";

        public static string NormalizeControlIdInGroup(string xmlText) {
            if (xmlText.Contains(Separator)) {
                throw new Exception($"`{Separator}` can not be used in resource xml");
            }
            var document = Parse(xmlText);
            if (document == null) {
                return "";
            }

            // get group controls, and enumerate it to normalize
            var groupNodes = document.GetElementsByTagName("group");
            if (groupNodes.Count == 0) {
                return document.ToXmlText();
            }

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

            // get menu root node, and enumerate it to normalize
            var menuRootNode = document.GetElementsByTagName("menu").OfType<XmlNode>().FirstOrDefault();
            if (menuRootNode == null) {
                return document.ToXmlText();
            }

            foreach (var childNode in menuRootNode.ChildNodes.OfType<XmlNode>()) {
                NormalizeControlIdWithXmlNode(childNode, menuId, null);
            }
            return document.ToXmlText();
        }

        private static void NormalizeControlIdWithXmlNode(XmlNode node, string suffix, Func<XmlNode, bool>? toBreak) {
            if (toBreak != null && toBreak.Invoke(node)) {
                return; // breakable, according to given condition
            }

            var attributes = node.Attributes;
            if (attributes != null) {
                var idAttribute = attributes["id"];
                var idValue = idAttribute?.Value;
                if (idAttribute != null && !string.IsNullOrWhiteSpace(idValue)) {
                    idAttribute.Value = $"{idValue!}{Separator}{suffix}"; // normalize id attribute
                }
            }

            foreach (var childNode in node.ChildNodes.OfType<XmlNode>()) {
                NormalizeControlIdWithXmlNode(childNode, suffix, toBreak);
            }
        }

        #endregion

        #region Template Related

        public static string ApplyTemplateForXml(string xmlText) {
            var document = Parse(xmlText);
            if (document == null) {
                return "";
            }

            // 1. find templates node from document
            var templatesNodes = document.GetElementsByTagName("__templates").OfType<XmlNode>().ToList();
            var subtreeTemplateNodes = document.SelectNodes("//*[@__as_subtree_template]")?.OfType<XmlNode>().ToList() ?? new List<XmlNode>();

            // 2.1. apply attribute template
            if (templatesNodes.Count != 0) {
                var attributeTemplates = ExtractAttributeTemplateDictionary(templatesNodes); // template name to attribute key-value pair
                ApplyAttributeTemplateForNodes(document, attributeTemplates);
            }

            // 2.2. apply subtree template
            var subtreeTemplates = ExtractSubtreeTemplateDictionary(templatesNodes, subtreeTemplateNodes); // template name to subtree node list
            HandleNestedSubtreeTemplate(document, subtreeTemplates);
            ApplySubtreeTemplateForNodes(document, subtreeTemplates);

            // 3. clear "__template" nodes and attributes
            foreach (var templatesNode in templatesNodes) {
                templatesNode.ParentNode?.RemoveChild(templatesNode);
            }
            foreach (var templateNode in subtreeTemplateNodes) {
                templateNode.Attributes?.RemoveNamedItem("__as_subtree_template");
            }

            // 4. returned the applied xml string
            return document.ToXmlText();
        }

        #region Attribute Template Related

        private static Dictionary<string, Dictionary<string, string>> ExtractAttributeTemplateDictionary(IEnumerable<XmlNode> templatesNodes) {
            var templateDictionary = new Dictionary<string, Dictionary<string, string>>();

            foreach (var templatesNode in templatesNodes) { // <__templates>
                var templateNodes = templatesNode.ChildNodes.OfType<XmlNode>().Where(n => n.Name == "__attribute_template");
                foreach (var templateNode in templateNodes) { // <__attribute_template>
                    var nodeAttributes = templateNode.Attributes;
                    var nameValue = nodeAttributes?["name"]?.Value.Trim();
                    if (string.IsNullOrWhiteSpace(nameValue)) {
                        continue;
                    }

                    var templateAttributes = new Dictionary<string, string>();
                    foreach (var attribute in nodeAttributes!.OfType<XmlAttribute>()) {
                        if (attribute.Name != "name") {
                            templateAttributes[attribute.Name] = attribute.Value;
                        }
                    }
                    templateDictionary[nameValue!] = templateAttributes;
                }
            }

            return templateDictionary;
        }

        private static void ApplyAttributeTemplateForNodes(XmlDocument document, Dictionary<string, Dictionary<string, string>> templateDictionary) {
            var nodesToBeApplied = document.SelectNodes("//*[@__template]")?.OfType<XmlNode>().ToList();
            if (nodesToBeApplied == null || nodesToBeApplied.Count == 0) {
                return;
            }
            foreach (var node in nodesToBeApplied) {
                ApplyAttributeTemplateForSingleNode(document, node, templateDictionary);
            }
        }

        private static void ApplyAttributeTemplateForSingleNode(XmlDocument document, XmlNode node, Dictionary<string, Dictionary<string, string>> templateDictionary) {
            var nodeAttributes = node.Attributes;
            var templateValue = nodeAttributes?["__template"]?.Value.Trim();
            if (templateValue == null || string.IsNullOrWhiteSpace(templateValue)) {
                return;
            }
            nodeAttributes!.RemoveNamedItem("__template"); // template attribute must be removed

            // get template name array
            var templateNames = templateValue.Split(',');
            if (templateNames.Length == 0) {
                return;
            }

            // apply each template name to node
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

        #endregion

        #region Subtree Template Related

        private static Dictionary<string, List<XmlNode>> ExtractSubtreeTemplateDictionary(IEnumerable<XmlNode> templatesNodes, IEnumerable<XmlNode> subtreeTemplateNodes) {
            var templateDictionary = new Dictionary<string, List<XmlNode>>();

            foreach (var templatesNode in templatesNodes) { // <__templates>
                var templateNodes = templatesNode.ChildNodes.OfType<XmlNode>().Where(n => n.Name == "__subtree_template");
                foreach (var templateNode in templateNodes) { // <__subtree_template>
                    var nodeAttributes = templateNode.Attributes;
                    var nameValue = nodeAttributes?["name"]?.Value.Trim();
                    if (string.IsNullOrWhiteSpace(nameValue)) {
                        continue;
                    }
                    templateDictionary[nameValue!] = templateNode.ChildNodes.OfType<XmlNode>().ToList();
                }
            }

            foreach (var templateNode in subtreeTemplateNodes) { // <xxx __as_subtree_template>
                var nodeAttributes = templateNode.Attributes;
                var nameValue = nodeAttributes?["__as_subtree_template"]?.Value.Trim();
                if (string.IsNullOrWhiteSpace(nameValue)) {
                    continue;
                }
                templateDictionary[nameValue!] = templateNode.ChildNodes.OfType<XmlNode>().ToList();
            }

            return templateDictionary;
        }

        private static void HandleNestedSubtreeTemplate(XmlDocument document, Dictionary<string, List<XmlNode>> templateDictionary) {
            var xmlns = new XmlNamespaceManager(document.NameTable);
            xmlns.AddNamespace("ns", "http://schemas.microsoft.com/office/2006/01/customui");

            // find nested templates
            var dependGraph = new Dictionary<string, HashSet<string>>();
            foreach (var templateKvPair in templateDictionary) {
                var (templateName, templateNodes) = (templateKvPair.Key, templateKvPair.Value);
                var dependNodes = templateNodes[0].ParentNode?.SelectNodes(".//ns:__use_template", xmlns)?.OfType<XmlNode>().ToList();
                if (dependNodes == null || dependNodes.Count == 0) {
                    dependGraph[templateName] = new HashSet<string>();
                } else {
                    dependGraph[templateName] = dependNodes
                        .Select(n => n.Attributes?["name"]?.Value)
                        .Where(n => !string.IsNullOrWhiteSpace(n))
                        .Cast<string>().ToHashSet();
                }
            }

            // handle nested templates in topological order
            while (dependGraph.Count != 0) {
                var freeTemplateNames = dependGraph
                    .Where(kv => kv.Value.Count == 0)
                    .Select(kv => kv.Key).ToList();
                if (freeTemplateNames.Count == 0) {
                    throw new Exception("There is dependency cycle in templates, which is not allowed");
                }

                foreach (var freeTemplateName in freeTemplateNames) {
                    var dependingTemplateNames = dependGraph
                        .Where(kv => kv.Value.Contains(freeTemplateName))
                        .Select(kv => kv.Key)
                        .ToArray(); // need to copy here, because of child removing

                    foreach (var templateName in dependingTemplateNames) {
                        var templateNodes = templateDictionary[templateName].ToList();
                        var parentNode = templateNodes[0].ParentNode;
                        if (parentNode == null) {
                            continue; // almost unreachable
                        }
                        var nodesToBeHandled = templateNodes[0].ParentNode?.SelectNodes($".//*[@name=\"{freeTemplateName}\"]")?.OfType<XmlNode>();
                        if (nodesToBeHandled != null) {
                            foreach (var node in nodesToBeHandled) { // pre-handle free template
                                ApplySubtreeTemplateForSingleNode(document, node, templateDictionary);
                            }
                        }
                        dependGraph[templateName].Remove(freeTemplateName);
                        templateDictionary[templateName] = parentNode.ChildNodes.OfType<XmlNode>().ToList(); // update template dictionary
                    }

                    dependGraph.Remove(freeTemplateName); // free template has been pre-handled
                }
            }
        }

        private static void ApplySubtreeTemplateForNodes(XmlDocument document, Dictionary<string, List<XmlNode>> templateDictionary) {
            var nodesToBeApplied = document.GetElementsByTagName("__use_template").OfType<XmlNode>().ToList();
            if (nodesToBeApplied.Count > 0) {
                foreach (var node in nodesToBeApplied) {
                    ApplySubtreeTemplateForSingleNode(document, node, templateDictionary);
                }
            }

            nodesToBeApplied = document.GetElementsByTagName("__use_reference").OfType<XmlNode>().ToList();
            if (nodesToBeApplied.Count > 0) {
                foreach (var node in nodesToBeApplied) {
                    ApplySubtreeTemplateForSingleNode(document, node, templateDictionary);
                }
            }
        }

        private static void ApplySubtreeTemplateForSingleNode(XmlDocument document, XmlNode node, IReadOnlyDictionary<string, List<XmlNode>> templateDictionary) {
            var templateName = node.Attributes?["name"]?.Value;
            if (string.IsNullOrWhiteSpace(templateName)) {
                return;
            }

            // get template subtree (node list) or reference (single node)
            List<XmlNode> templateNodeList;
            switch (node.Name) {
            case "__use_template":
                if (!templateDictionary.TryGetValue(templateName!, out templateNodeList)) {
                    return; // specific template is not found
                }
                break;
            case "__use_reference":
                var matched = new Regex(@"^\$(\w+?)=(\w+?)(?::(\d+?))?$").Match(templateName!);
                if (!matched.Success) {
                    return; // failed to extract referee
                }
                var (keyValue, valueValue, numValue) = (matched.Groups[1].Value, matched.Groups[2].Value, matched.Groups[3].Value);
                numValue = string.IsNullOrWhiteSpace(numValue) ? "1" : numValue!;
                if (string.IsNullOrWhiteSpace(keyValue) || string.IsNullOrWhiteSpace(valueValue) || !int.TryParse(numValue, out var num)) {
                    return; // use empty key or empty value or invalid num in template
                }
                var referenceNodes = document.SelectNodes($"//*[@{keyValue}=\"{valueValue}\"]")?.OfType<XmlNode>().ToList();
                if (referenceNodes == null || referenceNodes.Count == 0) {
                    return; // can not find template referee node
                }
                num = Math.Min(1, Math.Max(referenceNodes.Count, num));
                templateNodeList = new List<XmlNode> { referenceNodes[num - 1] };
                break;
            default:
                return;
            }

            // extract template rules, enumerate template node list, and apply to xml node
            var rules = SubtreeTemplateRules.ExtractFromXmlNode(node);
            foreach (var templateNode in templateNodeList) {
                var clonedNode = templateNode.CloneNode(true); // deep clone template node
                rules.ReplaceFieldValue(clonedNode);
                clonedNode = rules.RemoveSpecificNode(clonedNode);
                if (clonedNode != null) {
                    node.ParentNode?.InsertBefore(clonedNode, node);
                }
            }
            node.ParentNode?.RemoveChild(node); // subtree template node must be removed
        }


        class SubtreeTemplateRules {
            private readonly List<(string field, string from, string to, bool norec, bool re)> _replaceRules = new();
            private readonly List<(string field, string match, bool norec, bool re)> _removeRules = new();

            public static SubtreeTemplateRules ExtractFromXmlNode(XmlNode node) {
                var rulesObject = new SubtreeTemplateRules();

                var ruleNodes = node.ChildNodes.OfType<XmlNode>().ToList();
                ruleNodes.Add(node); // also regard the use_template or use_reference node as rule node

                foreach (var ruleNode in ruleNodes) {
                    var nodeName = ruleNode.Name;
                    var nodeAttributes = ruleNode.Attributes;

                    var fieldValue = nodeAttributes?["field"]?.Value;
                    var norecValue = nodeAttributes?["norec"]?.Value == "true";
                    if (nodeName == node.Name) { // allow to define rule on template node directly 
                        var replaceFieldValue = nodeAttributes?["replace_rule_field"]?.Value;
                        var removeFieldValue = nodeAttributes?["remove_rule_field"]?.Value;
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
                        var fromValue = nodeAttributes?["from"]?.Value;
                        var fromReValue = nodeAttributes?["from_re"]?.Value;
                        var toValue = nodeAttributes?["to"]?.Value;
                        if (string.IsNullOrWhiteSpace(fieldValue) || (string.IsNullOrWhiteSpace(fromValue) && string.IsNullOrWhiteSpace(fromReValue)) || toValue == null) {
                            continue;
                        }
                        if (!string.IsNullOrWhiteSpace(fromReValue) is var useReForReplacing && useReForReplacing) {
                            fromValue = fromReValue!; // use from_re instead
                        }
                        rulesObject._replaceRules.Add((fieldValue!, fromValue!, toValue, norecValue, useReForReplacing));
                        break;

                    case "__remove_rule":
                        var matchValue = nodeAttributes?["match"]?.Value;
                        var matchReValue = nodeAttributes?["match_re"]?.Value;
                        if (string.IsNullOrWhiteSpace(fieldValue) || (string.IsNullOrWhiteSpace(matchValue) && string.IsNullOrWhiteSpace(matchReValue))) {
                            continue;
                        }
                        if (!string.IsNullOrWhiteSpace(matchReValue) is var useReForRemoving && useReForRemoving) {
                            matchValue = matchReValue!; // use match_re instead
                        }
                        rulesObject._removeRules.Add((fieldValue!, matchValue!, norecValue, useReForRemoving));
                        break;

                    default:
                        continue;
                    }
                }

                return rulesObject;
            }

            public void ReplaceFieldValue(XmlNode node, int layer = 1) {
                if (_replaceRules.Count == 0) {
                    return;
                }

                var attributes = node.Attributes;
                foreach (var (field, from, to, norec, re) in _replaceRules) {
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
                    ReplaceFieldValue(childNode, layer + 1);
                }
            }

            public XmlNode? RemoveSpecificNode(XmlNode node, int layer = 1) {
                if (_removeRules.Count == 0) {
                    return node;
                }

                var attributes = node.Attributes;
                foreach (var (field, match, norec, re) in _removeRules) {
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

                var childNodes = node.ChildNodes.OfType<XmlNode>().ToArray(); // need to copy here, because of child removing
                foreach (var childNode in childNodes) {
                    var result = RemoveSpecificNode(childNode, layer + 1);
                    if (result == null) {
                        node.RemoveChild(childNode);
                    }
                }
                return node;
            }
        }

        #endregion

        #endregion

        #region Misc Methods

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
