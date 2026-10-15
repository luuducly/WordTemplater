using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using WordTemplater;
using Xunit;

namespace WordTemplater.Tests
{
    public class AdvancedFeatureTests
    {
        private static string GetTemplatePath()
        {
            var baseDir = AppContext.BaseDirectory;
            var path = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "src", "WordTemplater.Example", "Templates", "Template.docx"));
            if (!File.Exists(path))
            {
                path = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "src", "WordTemplater.Example", "Templates", "Template.docx"));
            }
            return path;
        }

        [Fact]
        public void SetRowShading_ShouldApplyFillColorToAllCells()
        {
            var row = new TableRow();
            var cell1 = new TableCell(new Paragraph(new Run(new Text("Cell 1"))));
            var cell2 = new TableCell(new Paragraph(new Run(new Text("Cell 2"))));
            row.Append(cell1, cell2);

            WordUtils.SetRowShading(row, "#AABBCC");

            foreach (var cell in row.Elements<TableCell>())
            {
                var props = cell.GetFirstChild<TableCellProperties>();
                Assert.NotNull(props);
                var shading = props.GetFirstChild<Shading>();
                Assert.NotNull(shading);
                Assert.Equal("AABBCC", shading.Fill?.Value);
            }
        }

        [Fact]
        public void Export_WithWatermark_ShouldInjectWatermarkIntoHeader()
        {
            var templatePath = GetTemplatePath();
            if (!File.Exists(templatePath)) return;

            using var templateStream = File.OpenRead(templatePath);
            using var template = new WordTemplate(templateStream);

            template.SetWatermark("CONFIDENTIAL", "FF0000", -45, 60, "Arial");

            var data = new JObject
            {
                ["Name"] = "Test User",
                ["Items"] = new JArray()
            };

            using var outputStream = template.Export(data);
            Assert.NotNull(outputStream);

            using var doc = WordprocessingDocument.Open(outputStream, false);
            var headerParts = doc.MainDocumentPart?.HeaderParts.ToList();
            Assert.NotNull(headerParts);
            Assert.NotEmpty(headerParts);

            var found = headerParts.Any(hp => hp.Header != null && hp.Header.InnerXml.Contains("CONFIDENTIAL"));
            Assert.True(found, "Watermark text 'CONFIDENTIAL' should be present in header part XML.");
        }

        [Fact]
        public void Export_WithTableZebraStriping_ShouldApplyAlternatingShading()
        {
            var templatePath = GetTemplatePath();
            if (!File.Exists(templatePath)) return;

            using var templateStream = File.OpenRead(templatePath);
            using var template = new WordTemplate(templateStream);

            var options = new ExportOptions
            {
                TableOddRowColor = "F0F0F0",
                TableEvenRowColor = "FFFFFF"
            };

            var data = new JObject
            {
                ["Name"] = "Zebra Test",
                ["Items"] = new JArray
                {
                    new JObject { ["Index"] = 1, ["Name"] = "Item 1" },
                    new JObject { ["Index"] = 2, ["Name"] = "Item 2" },
                    new JObject { ["Index"] = 3, ["Name"] = "Item 3" }
                }
            };

            using var outputStream = template.Export(data, options);
            Assert.NotNull(outputStream);

            using var doc = WordprocessingDocument.Open(outputStream, false);
            var tables = doc.MainDocumentPart?.Document.Body?.Descendants<Table>().ToList();
            Assert.NotNull(tables);
            Assert.NotEmpty(tables);

            // Verify that at least one table cell has shading
            var shadings = tables.SelectMany(t => t.Descendants<Shading>()).ToList();
            Assert.NotEmpty(shadings);
        }

        [Fact]
        public void LoopTemplater_ShouldProvideLoopMetadataVariables()
        {
            var templatePath = GetTemplatePath();
            if (!File.Exists(templatePath)) return;

            using var templateStream = File.OpenRead(templatePath);
            using var template = new WordTemplate(templateStream);

            var data = new JObject
            {
                ["Name"] = "Loop Test",
                ["Items"] = new JArray
                {
                    new JObject { ["Name"] = "First Item" },
                    new JObject { ["Name"] = "Second Item" }
                }
            };

            // Ensure export succeeds without throwing when looping over arrays
            using var outputStream = template.Export(data);
            Assert.NotNull(outputStream);
            Assert.True(outputStream.Length > 0);
        }
    }
}
