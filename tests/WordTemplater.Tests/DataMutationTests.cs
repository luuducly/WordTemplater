using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using WordTemplater;
using Xunit;
using static WordTemplater.Tests.TestTemplates;

namespace WordTemplater.Tests
{
    public class DataMutationTests
    {
        [Fact]
        public void Export_DoesNotMutateCallerData()
        {
            var data = Data(3);
            var before = data.ToString();

            using (Export(LoopIfTemplate(), data)) { }

            Assert.Equal(before, data.ToString());
        }

        [Fact]
        public void Export_DoesNotMutateCallerData_PrimitiveArray()
        {
            var template = Build(
                P(F("loop(Items)")),
                P(F(".")),
                P(F("endloop")));
            var data = new JObject { ["Items"] = new JArray("a", "b", "c") };
            var before = data.ToString();

            using (Export(template, data)) { }

            Assert.Equal(before, data.ToString());
        }

        [Fact]
        public void LoopMetadata_IsRendered()
        {
            using var output = Export(LoopMetadataTemplate(), Data(3));
            using var doc = WordprocessingDocument.Open(output, false);
            var paragraphs = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>()
                .Select(p => p.InnerText)
                .Where(t => t.Length > 0)
                .ToArray();

            Assert.Equal(new[]
            {
                "1|True|False|False|True|item 0",
                "2|False|False|True|False|item 1",
                "3|False|True|False|True|item 2",
            }, paragraphs);
        }

        private static Stream Export(byte[] templateBytes, JObject data)
        {
            using var templateStream = new MemoryStream(templateBytes);
            using var template = new WordTemplate(templateStream);
            var output = template.Export(data);
            Assert.NotNull(output);
            return output!;
        }
    }
}
