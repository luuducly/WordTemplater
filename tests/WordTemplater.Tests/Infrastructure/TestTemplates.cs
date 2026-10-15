using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;

namespace WordTemplater.Tests
{
    /// <summary>
    /// Programmatic templates and data builders shared by the golden, lifetime, concurrency and mutation tests.
    /// </summary>
    internal static class TestTemplates
    {
        /// <summary>
        /// A merge field the library recognizes (the code must contain "  \* MERGEFORMAT" with two spaces).
        /// </summary>
        public static SimpleField F(string code) =>
            new SimpleField(new Run(new Text("«" + code + "»"))) { Instruction = " MERGEFIELD  " + code + "  \\* MERGEFORMAT " };

        public static Paragraph P(params OpenXmlElement[] children) => new Paragraph(children);

        public static Run T(string text) => new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        /// <summary>
        /// Builds a .docx whose body contains the given top-level elements.
        /// </summary>
        public static byte[] Build(params OpenXmlElement[] bodyChildren)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body(bodyChildren));
            }
            return ms.ToArray();
        }

        /// <summary>
        /// Title, then table(Items) over a table with a plain header row and one template row, then endtable.
        /// </summary>
        public static byte[] TableTemplate(string? extraCellCode = null)
        {
            var header = new TableRow(
                new TableCell(P(T("No"))),
                new TableCell(P(T("Name"))),
                new TableCell(P(T("Qty"))),
                new TableCell(P(T("Price"))));
            var row = new TableRow(
                new TableCell(P(F("_index"))),
                new TableCell(P(F("Name:upper()"))),
                new TableCell(P(F("Qty:number(#,##0)"))),
                new TableCell(P(F("Price:currency(en-US)"))));
            if (extraCellCode != null)
            {
                row.Append(new TableCell(P(F(extraCellCode))));
            }
            return Build(
                P(F("Title")),
                P(F("table(Items)")),
                new Table(header, row),
                P(F("endtable")),
                P(T("end")));
        }

        /// <summary>
        /// Like <see cref="TableTemplate"/>, but the table also has a plain-text footer row after the template row,
        /// so the repeated rows are inserted before an existing row.
        /// </summary>
        public static byte[] TableWithFooterTemplate()
        {
            var header = new TableRow(
                new TableCell(P(T("No"))),
                new TableCell(P(T("Name"))),
                new TableCell(P(T("Qty"))),
                new TableCell(P(T("Price"))));
            var row = new TableRow(
                new TableCell(P(F("_index"))),
                new TableCell(P(F("Name:upper()"))),
                new TableCell(P(F("Qty:number(#,##0)"))),
                new TableCell(P(F("Price:currency(en-US)"))));
            var footer = new TableRow(
                new TableCell(P(T("Total"))),
                new TableCell(P()),
                new TableCell(P()),
                new TableCell(P()));
            return Build(
                P(F("Title")),
                P(F("table(Items)")),
                new Table(header, row, footer),
                P(F("endtable")),
                P(T("end")));
        }

        /// <summary>
        /// Paragraph loop over Groups containing a paragraph loop over Items with a nested if(Qty &gt; 5) block.
        /// </summary>
        public static byte[] NestedLoopTemplate()
        {
            return Build(
                P(F("Title")),
                P(F("loop(Groups)")),
                P(F("Name")),
                P(F("loop(Items)")),
                P(F("_index"), T(". "), F("Name")),
                P(F("if(Qty > 5)")),
                P(T("big: "), F("Qty")),
                P(F("endif")),
                P(F("endloop")),
                P(T("after inner")),
                P(F("endloop")),
                P(T("end")));
        }

        /// <summary>
        /// Paragraph loop over Items with a nested if(Qty &gt; 5) block.
        /// </summary>
        public static byte[] LoopIfTemplate()
        {
            return Build(
                P(F("Title")),
                P(F("loop(Items)")),
                P(F("_index"), T(". "), F("Name")),
                P(F("if(Qty > 5)")),
                P(T("big qty: "), F("Qty")),
                P(F("endif")),
                P(F("endloop")),
                P(T("end")));
        }

        /// <summary>
        /// Paragraph loop over Items rendering the loop metadata fields separated by "|".
        /// </summary>
        public static byte[] LoopMetadataTemplate()
        {
            return Build(
                P(F("loop(Items)")),
                P(F("_index"), T("|"), F("_first"), T("|"), F("_last"), T("|"), F("_even"), T("|"), F("_odd"), T("|"), F("Name")),
                P(F("endloop")));
        }

        public static JObject Data(int n, bool withCode = false)
        {
            var items = new JArray();
            for (int i = 0; i < n; i++)
            {
                var item = new JObject
                {
                    ["Name"] = "item " + i,
                    ["Qty"] = i % 10,
                    ["Price"] = i * 1.5m
                };
                if (withCode)
                {
                    item["Code"] = "CODE-" + i;
                }
                items.Add(item);
            }
            return new JObject
            {
                ["Title"] = "Report",
                ["Items"] = items
            };
        }

        /// <summary>
        /// Data for <see cref="NestedLoopTemplate"/>: Groups[g] = { Name, Items[itemsPerGroup] }; group 1 (if any) has no items.
        /// </summary>
        public static JObject NestedData(int groups, int itemsPerGroup)
        {
            var groupArray = new JArray();
            for (int g = 0; g < groups; g++)
            {
                var items = new JArray();
                if (g != 1)
                {
                    for (int i = 0; i < itemsPerGroup; i++)
                    {
                        var n = g * itemsPerGroup + i;
                        items.Add(new JObject
                        {
                            ["Name"] = "item " + g + "." + i,
                            ["Qty"] = n % 10,
                            ["Price"] = n * 1.5m
                        });
                    }
                }
                groupArray.Add(new JObject
                {
                    ["Name"] = "group " + g,
                    ["Items"] = items
                });
            }
            return new JObject
            {
                ["Title"] = "Report",
                ["Groups"] = groupArray
            };
        }
    }
}
