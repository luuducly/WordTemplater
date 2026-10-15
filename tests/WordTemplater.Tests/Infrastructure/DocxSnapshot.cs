using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace WordTemplater.Tests
{
    /// <summary>
    /// Turns an exported .docx into a deterministic, human-diffable text snapshot.
    /// Volatile values (relationship ids, drawing ids/names, edit ids) are replaced by placeholders numbered
    /// in order of first appearance, so two exports of the same input normalize to the same text.
    /// </summary>
    internal static class DocxSnapshot
    {
        private const string RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string Wp14Ns = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
        private static readonly XName WpDocPr = XName.Get("docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        private static readonly XName PicCNvPr = XName.Get("cNvPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");

        public static string Normalize(Stream docx)
        {
            docx.Position = 0;
            using var doc = WordprocessingDocument.Open(docx, false);
            var main = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
            var state = new State();
            var sb = new StringBuilder();

            var xmlParts = new List<OpenXmlPart> { main };
            xmlParts.AddRange(main.Parts.Select(p => p.OpenXmlPart).Where(p => p is HeaderPart || p is FooterPart));

            foreach (var part in xmlParts)
            {
                XDocument xml;
                using (var s = part.GetStream(FileMode.Open, FileAccess.Read))
                {
                    xml = XDocument.Load(s, LoadOptions.PreserveWhitespace);
                }
                ReplaceVolatileValues(xml, part, state);
                sb.Append("=== ").Append(part.Uri).Append(" ===\n");
                sb.Append(Serialize(xml)).Append('\n');
            }

            sb.Append("=== relationships ===\n");
            foreach (var rel in state.Relationships)
            {
                sb.Append(rel.Placeholder).Append(" -> ").Append(DescribeRelationship(rel.Source, rel.Id)).Append('\n');
            }

            sb.Append("=== parts ===\n");
            foreach (var part in AllParts(doc).OrderBy(p => p.Uri.OriginalString, StringComparer.Ordinal))
            {
                sb.Append(part.Uri).Append(' ').Append(part.ContentType);
                if (!IsVolatileXmlPart(part))
                {
                    sb.Append(' ').Append(Hash(part));
                }
                sb.Append('\n');
            }

            return sb.ToString();
        }

        private static void ReplaceVolatileValues(XDocument xml, OpenXmlPart part, State state)
        {
            foreach (var element in xml.Descendants())
            {
                bool isDrawingProps = element.Name == WpDocPr || element.Name == PicCNvPr;
                foreach (var attr in element.Attributes())
                {
                    if (attr.IsNamespaceDeclaration) continue;
                    var name = attr.Name;
                    if (name.NamespaceName == RelationshipsNs)
                    {
                        attr.Value = state.Relationship(part, attr.Value);
                    }
                    else if (name.NamespaceName == Wp14Ns && (name.LocalName == "editId" || name.LocalName == "anchorId"))
                    {
                        attr.Value = state.Map("EDITID", attr.Value);
                    }
                    else if (isDrawingProps && name.Namespace == XNamespace.None && name.LocalName == "id")
                    {
                        attr.Value = state.Map("DOCPR_ID", attr.Value);
                    }
                    else if (isDrawingProps && name.Namespace == XNamespace.None && name.LocalName == "name")
                    {
                        attr.Value = state.Map("DOCPR_NAME", attr.Value);
                    }
                }
            }
        }

        private static string Serialize(XDocument xml)
        {
            var settings = new XmlWriterSettings
            {
                Indent = true,
                OmitXmlDeclaration = true,
                NewLineChars = "\n"
            };
            var sb = new StringBuilder();
            using (var writer = XmlWriter.Create(sb, settings))
            {
                xml.Save(writer);
            }
            return sb.ToString();
        }

        private static string DescribeRelationship(OpenXmlPart source, string id)
        {
            foreach (var pair in source.Parts)
            {
                if (pair.RelationshipId != id) continue;
                var internalPart = pair.OpenXmlPart;
                return IsVolatileXmlPart(internalPart)
                    ? internalPart.ContentType
                    : internalPart.ContentType + " " + Hash(internalPart);
            }
            var hyperlink = source.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
            if (hyperlink != null) return "hyperlink " + hyperlink.Uri.OriginalString;
            var external = source.ExternalRelationships.FirstOrDefault(r => r.Id == id);
            if (external != null) return "external " + external.Uri.OriginalString;
            var dataPart = source.DataPartReferenceRelationships.FirstOrDefault(r => r.Id == id);
            if (dataPart != null) return "data " + dataPart.DataPart.ContentType + " " + HashBytes(ReadAll(dataPart.DataPart.GetStream()));
            return "unresolved";
        }

        /// <summary>
        /// Main/header/footer parts embed volatile ids, so they are snapshotted as normalized XML instead of hashed.
        /// </summary>
        private static bool IsVolatileXmlPart(OpenXmlPart part) =>
            part is MainDocumentPart || part is HeaderPart || part is FooterPart;

        private static IEnumerable<OpenXmlPart> AllParts(WordprocessingDocument doc)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var stack = new Stack<OpenXmlPart>(doc.Parts.Select(p => p.OpenXmlPart));
            while (stack.Count > 0)
            {
                var part = stack.Pop();
                if (!seen.Add(part.Uri.OriginalString)) continue;
                yield return part;
                foreach (var child in part.Parts)
                {
                    stack.Push(child.OpenXmlPart);
                }
            }
        }

        /// <summary>
        /// First 16 hex chars of SHA-256 over the part content. XML parts are hashed in a canonical form
        /// (parsed and re-serialized without declaration) so that a mere re-save by the SDK does not count as a change.
        /// </summary>
        private static string Hash(OpenXmlPart part)
        {
            byte[] bytes;
            using (var s = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                bytes = ReadAll(s);
            }
            var contentType = part.ContentType;
            if (contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) || contentType.EndsWith("/xml", StringComparison.OrdinalIgnoreCase))
            {
                using var ms = new MemoryStream(bytes);
                var xml = XDocument.Load(ms, LoadOptions.PreserveWhitespace);
                bytes = Encoding.UTF8.GetBytes(xml.Root!.ToString(SaveOptions.DisableFormatting));
            }
            return HashBytes(bytes);
        }

        private static string HashBytes(byte[] bytes) =>
            Convert.ToHexString(SHA256.HashData(bytes)).Substring(0, 16).ToLowerInvariant();

        private static byte[] ReadAll(Stream s)
        {
            using (s)
            using (var ms = new MemoryStream())
            {
                s.CopyTo(ms);
                return ms.ToArray();
            }
        }

        private sealed class State
        {
            private readonly Dictionary<string, Dictionary<string, string>> _maps = new Dictionary<string, Dictionary<string, string>>();
            private readonly Dictionary<(string Source, string Id), string> _relationshipMap = new Dictionary<(string, string), string>();

            public List<(string Placeholder, OpenXmlPart Source, string Id)> Relationships { get; } = new List<(string, OpenXmlPart, string)>();

            public string Map(string kind, string value)
            {
                if (!_maps.TryGetValue(kind, out var map))
                {
                    map = new Dictionary<string, string>(StringComparer.Ordinal);
                    _maps[kind] = map;
                }
                if (!map.TryGetValue(value, out var placeholder))
                {
                    placeholder = kind + (map.Count + 1);
                    map[value] = placeholder;
                }
                return placeholder;
            }

            /// <summary>
            /// Relationship ids are scoped to their source part, so the key is (source part, id).
            /// </summary>
            public string Relationship(OpenXmlPart source, string id)
            {
                var key = (source.Uri.OriginalString, id);
                if (!_relationshipMap.TryGetValue(key, out var placeholder))
                {
                    placeholder = "REL" + (_relationshipMap.Count + 1);
                    _relationshipMap[key] = placeholder;
                    Relationships.Add((placeholder, source, id));
                }
                return placeholder;
            }
        }
    }
}
