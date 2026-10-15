using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;
using WordTemplater;
using Xunit;

namespace WordTemplater.Tests
{
    /// <summary>
    /// Golden/snapshot tests: the normalized export (see <see cref="DocxSnapshot"/>) must match
    /// Snapshots/{name}.approved.txt. A missing approved file is created (and the test fails once);
    /// a mismatch writes Snapshots/{name}.received.txt for diffing.
    /// </summary>
    public class GoldenOutputTests
    {
        [Fact]
        public void TemplateOrigin()
        {
            var template = File.ReadAllBytes(TestPaths.Template("TemplateOrigin.docx"));
            VerifySnapshot("TemplateOrigin", template, () =>
            {
                var data = JObject.Parse(File.ReadAllText(TestPaths.DataSample("DataOrigin.json")));
                data["Word"] = Convert.ToBase64String(File.ReadAllBytes(TestPaths.Template("Equation.docx")));
                data["Image"] = Convert.ToBase64String(File.ReadAllBytes(TestPaths.Template("Author.jpg")));
                data["RectImage"] = Convert.ToBase64String(File.ReadAllBytes(TestPaths.Template("Rectangle.png")));
                return data;
            }, configure: t =>
            {
                t.RegisterEvaluator("customizable", new PrefixEvaluator());
                t.RegisterEvaluator("upperFirstLetter", new UpperFirstLetterEvaluator());
            });
        }

        [Fact]
        public void Template()
        {
            var template = File.ReadAllBytes(TestPaths.Template("Template.docx"));
            VerifySnapshot("Template", template, () =>
            {
                var data = JObject.Parse(File.ReadAllText(TestPaths.DataSample("Data.json")));
                // "image" holds Mermaid sources rendered over the network (kroki.io); the tests must stay offline.
                RemovePropertiesNamed(data, "image");
                Assert.DoesNotContain(data.Descendants().OfType<JProperty>(), p => string.Equals(p.Name, "image", StringComparison.OrdinalIgnoreCase));
                return data;
            });
        }

        [Fact]
        public void Table20Zebra()
        {
            VerifySnapshot("Table20Zebra", TestTemplates.TableTemplate(), () => TestTemplates.Data(20),
                options: new ExportOptions { TableOddRowColor = "F0F0F0", TableEvenRowColor = "FFFFFF" });
        }

        [Fact]
        public void LoopIf12()
        {
            VerifySnapshot("LoopIf12", TestTemplates.LoopIfTemplate(), () => TestTemplates.Data(12));
        }

        [Fact]
        public void BarcodeQr5_Barcode()
        {
            VerifySnapshot("BarcodeQr5_Barcode", TestTemplates.TableTemplate("Code:barcode()"), () => TestTemplates.Data(5, true));
        }

        [Fact]
        public void BarcodeQr5_QrCode()
        {
            VerifySnapshot("BarcodeQr5_QrCode", TestTemplates.TableTemplate("Code:qrcode()"), () => TestTemplates.Data(5, true));
        }

        [Fact]
        public void Empty_Table()
        {
            VerifySnapshot("Empty_Table", TestTemplates.TableTemplate(), () => TestTemplates.Data(0));
        }

        [Fact]
        public void Empty_LoopIf()
        {
            VerifySnapshot("Empty_LoopIf", TestTemplates.LoopIfTemplate(), () => TestTemplates.Data(0));
        }

        [Fact]
        public void TableWithFooter7()
        {
            VerifySnapshot("TableWithFooter7", TestTemplates.TableWithFooterTemplate(), () => TestTemplates.Data(7),
                options: new ExportOptions { TableOddRowColor = "F0F0F0", TableEvenRowColor = "FFFFFF" });
        }

        [Fact]
        public void NestedLoops4x3()
        {
            VerifySnapshot("NestedLoops4x3", TestTemplates.NestedLoopTemplate(), () => TestTemplates.NestedData(4, 3));
        }

        private static void VerifySnapshot(string name, byte[] template, Func<JObject> data,
            Action<WordTemplate>? configure = null, ExportOptions? options = null)
        {
            var previousCulture = CultureInfo.CurrentCulture;
            var previousUICulture = CultureInfo.CurrentUICulture;
            string actual, again;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
                CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("en-US");
                // Two independent exports must normalize identically (process-wide id counter, GUIDs, random edit ids).
                actual = ExportAndNormalize(template, data(), configure, options);
                again = ExportAndNormalize(template, data(), configure, options);
            }
            finally
            {
                CultureInfo.CurrentCulture = previousCulture;
                CultureInfo.CurrentUICulture = previousUICulture;
            }

            if (actual != again)
            {
                var (line, first, second) = FirstDifference(actual, again);
                Assert.Fail($"Snapshot '{name}' is not deterministic: two exports of the same input differ at line {line}.\n first : {first}\n second: {second}");
            }

            Directory.CreateDirectory(TestPaths.SnapshotsDir);
            var approvedPath = Path.Combine(TestPaths.SnapshotsDir, name + ".approved.txt");
            var receivedPath = Path.Combine(TestPaths.SnapshotsDir, name + ".received.txt");

            if (!File.Exists(approvedPath))
            {
                File.WriteAllText(approvedPath, actual);
                Assert.Fail($"Approved snapshot was missing and has been created: {approvedPath}. Review it and re-run the test.");
            }

            var approved = File.ReadAllText(approvedPath).Replace("\r\n", "\n");
            if (approved == actual)
            {
                if (File.Exists(receivedPath)) File.Delete(receivedPath);
                return;
            }

            File.WriteAllText(receivedPath, actual);
            var (diffLine, expectedLine, actualLine) = FirstDifference(approved, actual);
            Assert.Fail($"Snapshot '{name}' differs from {approvedPath} at line {diffLine}.\n approved: {expectedLine}\n received: {actualLine}\nFull output written to {receivedPath}");
        }

        private static string ExportAndNormalize(byte[] templateBytes, JObject data, Action<WordTemplate>? configure, ExportOptions? options)
        {
            using var templateStream = new MemoryStream(templateBytes);
            using var template = new WordTemplate(templateStream);
            configure?.Invoke(template);
            using var output = options == null ? template.Export(data) : template.Export(data, options);
            Assert.NotNull(output);
            return DocxSnapshot.Normalize(output!);
        }

        private static (int Line, string Expected, string Actual) FirstDifference(string expected, string actual)
        {
            var e = expected.Split('\n');
            var a = actual.Split('\n');
            for (int i = 0; i < Math.Max(e.Length, a.Length); i++)
            {
                var el = i < e.Length ? e[i] : "<end of file>";
                var al = i < a.Length ? a[i] : "<end of file>";
                if (el != al) return (i + 1, el, al);
            }
            return (0, string.Empty, string.Empty);
        }

        private static void RemovePropertiesNamed(JToken token, string name)
        {
            if (token is JObject obj)
            {
                foreach (var property in obj.Properties().Where(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)).ToList())
                {
                    property.Remove();
                }
            }
            foreach (var child in token.Children().ToList())
            {
                RemovePropertiesNamed(child, name);
            }
        }

        /// <summary>
        /// Deterministic stand-in for the example's Number2Text evaluator.
        /// </summary>
        private sealed class PrefixEvaluator : IEvaluator
        {
            public string Evaluate(object fieldValue, List<object> parameters) => "N2T:" + fieldValue;
        }

        private sealed class UpperFirstLetterEvaluator : IEvaluator
        {
            public string Evaluate(object fieldValue, List<object> parameters)
            {
                var s = fieldValue?.ToString() ?? string.Empty;
                return s.Length == 0 ? s : char.ToUpperInvariant(s[0]) + s.Substring(1);
            }
        }
    }
}
