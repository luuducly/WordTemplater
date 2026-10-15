using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using WordTemplater;
using Xunit;

namespace WordTemplater.Tests
{
    public class NewFeaturesTests
    {
        [Fact]
        public void DateEvaluator_ShouldFormatDatesCorrectly()
        {
            var evaluator = new DateEvaluator();

            // Date string ISO
            var res1 = evaluator.Evaluate("2026-10-15T14:30:00", new List<object> { "dd/MM/yyyy" });
            Assert.Equal("15/10/2026", res1);

            // DateTime object
            var dt = new DateTime(2026, 12, 25, 9, 0, 0);
            var res2 = evaluator.Evaluate(dt, new List<object> { "yyyy-MM-dd HH:mm" });
            Assert.Equal("2026-12-25 09:00", res2);

            // Culture support
            var res3 = evaluator.Evaluate(new DateTime(2026, 1, 1), new List<object> { "MMMM", "vi-VN" });
            Assert.Contains("tháng", res3.ToLower());
        }

        [Fact]
        public void NumberEvaluator_ShouldFormatNumbersWithCulture()
        {
            var evaluator = new NumberEvaluator();

            // Decimal / double
            var res1 = evaluator.Evaluate(1234567.89, new List<object> { "#,##0.00", "en-US" });
            Assert.Equal("1,234,567.89", res1);

            // String from JSON
            var res2 = evaluator.Evaluate("9876543", new List<object> { "#,##0", "vi-VN" });
            Assert.Equal("9.876.543", res2);
        }

        [Fact]
        public void DefaultValueEvaluator_ShouldFallbackWhenValueIsNullorEmpty()
        {
            var evaluator = new DefaultValueEvaluator();

            var resNull = evaluator.Evaluate(null!, new List<object> { "N/A" });
            Assert.Equal("N/A", resNull);

            var resEmpty = evaluator.Evaluate("", new List<object> { "DefaultVal" });
            Assert.Equal("DefaultVal", resEmpty);

            var resValid = evaluator.Evaluate("ActualValue", new List<object> { "DefaultVal" });
            Assert.Equal("ActualValue", resValid);
        }

        [Fact]
        public void PaddingEvaluators_ShouldPadLeftAndRight()
        {
            var padLeft = new PadLeftEvaluator();
            var resLeft = padLeft.Evaluate(42, new List<object> { 5, "0" });
            Assert.Equal("00042", resLeft);

            var padRight = new PadRightEvaluator();
            var resRight = padRight.Evaluate("INV", new List<object> { 6, "-" });
            Assert.Equal("INV---", resRight);
        }

        [Fact]
        public void ResolveToken_ShouldResolveNestedAndArrayProperties()
        {
            var json = @"{
                ""company"": {
                    ""name"": ""Google LLC"",
                    ""address"": {
                        ""city"": ""Mountain View"",
                        ""zip"": 94043
                    }
                },
                ""employees"": [
                    { ""name"": ""John Doe"", ""role"": ""Engineer"" },
                    { ""name"": ""Jane Smith"", ""role"": ""Architect"" }
                ]
            }";
            var data = JObject.Parse(json);

            // Direct property
            Assert.NotNull(Utils.ResolveToken(data, "company"));

            // Dot notation
            var cityToken = Utils.ResolveToken(data, "company.address.city");
            Assert.NotNull(cityToken);
            Assert.Equal("Mountain View", cityToken.ToString());

            // Case-insensitive dot notation
            var zipToken = Utils.ResolveToken(data, "Company.Address.Zip");
            Assert.NotNull(zipToken);
            Assert.Equal("94043", zipToken.ToString());

            // Array indexing
            var empName = Utils.ResolveToken(data, "employees[0].name");
            Assert.NotNull(empName);
            Assert.Equal("John Doe", empName.ToString());

            // Non-existent path returns null
            Assert.Null(Utils.ResolveToken(data, "company.missing.path"));
        }

        [Fact]
        public void IfEvaluator_InlineTernary_TruthinessAndComparisons()
        {
            var evaluator = new IfEvaluator();

            // 1 parameter: truthy check (return param or empty)
            Assert.Equal("YES", evaluator.Evaluate(true, new List<object> { "YES" }));
            Assert.Equal(string.Empty, evaluator.Evaluate(false, new List<object> { "YES" }));

            // 2 parameters: truthy check (ternary trueBranch, falseBranch)
            Assert.Equal("APPROVED", evaluator.Evaluate(true, new List<object> { "APPROVED", "DENIED" }));
            Assert.Equal("DENIED", evaluator.Evaluate(false, new List<object> { "APPROVED", "DENIED" }));
            Assert.Equal("APPROVED", evaluator.Evaluate("true", new List<object> { "APPROVED", "DENIED" }));
            Assert.Equal("DENIED", evaluator.Evaluate("false", new List<object> { "APPROVED", "DENIED" }));
            Assert.Equal("APPROVED", evaluator.Evaluate(1, new List<object> { "APPROVED", "DENIED" }));
            Assert.Equal("DENIED", evaluator.Evaluate(0, new List<object> { "APPROVED", "DENIED" }));

            // 3 parameters: equality check (field == compare ? v1 : v2)
            Assert.Equal("Gold Tier", evaluator.Evaluate("VIP", new List<object> { "VIP", "Gold Tier", "Standard Tier" }));
            Assert.Equal("Standard Tier", evaluator.Evaluate("BASIC", new List<object> { "VIP", "Gold Tier", "Standard Tier" }));

            // 4 parameters: comparison operators (op, compareValue, v1, v2)
            Assert.Equal("PASSED", evaluator.Evaluate(85, new List<object> { ">=", 80, "PASSED", "FAILED" }));
            Assert.Equal("FAILED", evaluator.Evaluate(75, new List<object> { ">=", 80, "PASSED", "FAILED" }));
            Assert.Equal("MAX", evaluator.Evaluate(100, new List<object> { "==", 100, "MAX", "NOT MAX" }));
            Assert.Equal("NOT EQUAL", evaluator.Evaluate("admin", new List<object> { "!=", "guest", "NOT EQUAL", "EQUAL" }));
        }

        [Fact]
        public void BarCodeAndQRCode_GenerateValidPngBytes()
        {
            var qrBytes = Utils.GetQRCodeImageBytes("https://github.com/luuducly/WordTemplater");
            Assert.NotNull(qrBytes);
            Assert.True(qrBytes.Length > 0);

            var barcodeBytes = Utils.GetBarCodeImageBytes("PROD-998877");
            Assert.NotNull(barcodeBytes);
            Assert.True(barcodeBytes.Length > 0);
        }

        [Fact]
        public void Export_FullRichMediaAndConditional_EmbedsQRBarcodeAndEvaluatesIf()
        {
            var templateBytes = TestTemplates.Build(
                TestTemplates.P(TestTemplates.F("qrcode(QrUrl)")),
                TestTemplates.P(TestTemplates.F("barcode(ProductSku)")),
                TestTemplates.P(TestTemplates.F("if(IsActive, 'ACTIVE_USER', 'DISABLED_USER')")),
                TestTemplates.P(TestTemplates.F("if(Score >= 80, 'HIGH_HONORS', 'STANDARD')"))
            );

            using var templateStream = new MemoryStream(templateBytes);
            using var template = new WordTemplate(templateStream);

            var data = new JObject
            {
                ["QrUrl"] = "https://github.com/luuducly/WordTemplater",
                ["ProductSku"] = "SKU-998877",
                ["IsActive"] = true,
                ["Score"] = 92
            };

            using var output = template.Export(data);
            Assert.NotNull(output);

            using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(output!, false);
            var imageParts = doc.MainDocumentPart!.ImageParts.ToList();
            Assert.Equal(2, imageParts.Count);

            var bodyText = doc.MainDocumentPart.Document.Body!.InnerText;
            Assert.Contains("ACTIVE_USER", bodyText);
            Assert.DoesNotContain("DISABLED_USER", bodyText);
            Assert.Contains("HIGH_HONORS", bodyText);
            Assert.DoesNotContain("STANDARD", bodyText);
        }

        [Fact]
        public void Export_FullShowcaseTemplate_GeneratesValidDocument()
        {
            var templateBytes = TestTemplates.Build(
                TestTemplates.P(TestTemplates.T("WORDTEMPLATER SHOWCASE")),
                TestTemplates.P(TestTemplates.F("image(Logo)")),
                TestTemplates.P(TestTemplates.F("qrcode(QrUrl)")),
                TestTemplates.P(TestTemplates.F("barcode(Sku)")),
                TestTemplates.P(TestTemplates.F("if(IsActive, 'ACTIVE', 'INACTIVE')")),
                TestTemplates.P(TestTemplates.F("if(Score >= 80, 'HONORS', 'PASS')")),
                TestTemplates.P(TestTemplates.F("padleft(Id, 6, '0')")),
                TestTemplates.P(TestTemplates.F("date(Created, 'dd/MM/yyyy')")),
                TestTemplates.P(TestTemplates.F("currency(Total, 'en-US')"))
            );

            using var templateStream = new MemoryStream(templateBytes);
            using var template = new WordTemplate(templateStream);

            var data = new JObject
            {
                ["Logo"] = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
                ["QrUrl"] = "https://github.com/luuducly/WordTemplater",
                ["Sku"] = "SKU-9988",
                ["IsActive"] = true,
                ["Score"] = 90,
                ["Id"] = 42,
                ["Created"] = "2026-10-15T12:00:00",
                ["Total"] = 1500.50
            };

            using var output = template.Export(data);
            Assert.NotNull(output);
            using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(output!, false);
            var text = doc.MainDocumentPart!.Document.Body!.InnerText;
            Assert.Contains("ACTIVE", text);
            Assert.Contains("HONORS", text);
            Assert.Contains("000042", text);
            Assert.Contains("15/10/2026", text);
            Assert.Contains("$1,500.50", text);
        }

        [Fact]
        public void Export_PreprintedBillWithTextShapes_HeaderBodyFooter_FillsExactPositions()
        {
            using var ms = new MemoryStream();
            using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();

                // 1. Header with Text Shape
                var headerPart = main.AddNewPart<DocumentFormat.OpenXml.Packaging.HeaderPart>();
                var headerShape = new DocumentFormat.OpenXml.Vml.Shape
                {
                    Id = "H_Shape",
                    Style = "position:absolute;margin-left:350pt;margin-top:20pt;width:150pt;height:20pt",
                    Filled = false,
                    Stroked = false
                };
                var hTxbx = new DocumentFormat.OpenXml.Vml.TextBox();
                var hTxbxContent = new DocumentFormat.OpenXml.Wordprocessing.TextBoxContent(TestTemplates.P(TestTemplates.F("InvoiceSerial")));
                hTxbx.Append(hTxbxContent);
                headerShape.Append(hTxbx);
                headerPart.Header = new DocumentFormat.OpenXml.Wordprocessing.Header(
                    TestTemplates.P(TestTemplates.T("Ký hiệu hóa đơn: .................")),
                    TestTemplates.P(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Picture(headerShape)))
                );

                // 2. Footer with Text Shape
                var footerPart = main.AddNewPart<DocumentFormat.OpenXml.Packaging.FooterPart>();
                var footerShape = new DocumentFormat.OpenXml.Vml.Shape
                {
                    Id = "F_Shape",
                    Style = "position:absolute;margin-left:300pt;margin-top:10pt;width:180pt;height:20pt",
                    Filled = false,
                    Stroked = false
                };
                var fTxbx = new DocumentFormat.OpenXml.Vml.TextBox();
                var fTxbxContent = new DocumentFormat.OpenXml.Wordprocessing.TextBoxContent(TestTemplates.P(TestTemplates.F("SignerName:upper()")));
                fTxbx.Append(fTxbxContent);
                footerShape.Append(fTxbx);
                footerPart.Footer = new DocumentFormat.OpenXml.Wordprocessing.Footer(
                    TestTemplates.P(TestTemplates.T("Người ký duyệt: .................")),
                    TestTemplates.P(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Picture(footerShape)))
                );

                // 3. Body with Background dots and Overlay Text Shape
                var bodyShape = new DocumentFormat.OpenXml.Vml.Shape
                {
                    Id = "B_Shape",
                    Style = "position:absolute;margin-left:140pt;margin-top:15pt;width:260pt;height:20pt",
                    Filled = false,
                    Stroked = false
                };
                var bTxbx = new DocumentFormat.OpenXml.Vml.TextBox();
                var bTxbxContent = new DocumentFormat.OpenXml.Wordprocessing.TextBoxContent(TestTemplates.P(TestTemplates.F("CustomerName")));
                bTxbx.Append(bTxbxContent);
                bodyShape.Append(bTxbx);

                var sectionProps = new DocumentFormat.OpenXml.Wordprocessing.SectionProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.HeaderReference { Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default, Id = main.GetIdOfPart(headerPart) },
                    new DocumentFormat.OpenXml.Wordprocessing.FooterReference { Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default, Id = main.GetIdOfPart(footerPart) }
                );

                main.Document.Body = new DocumentFormat.OpenXml.Wordprocessing.Body(
                    TestTemplates.P(TestTemplates.T("Họ tên khách hàng: ................................................................")),
                    TestTemplates.P(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Picture(bodyShape))),
                    sectionProps
                );
            }

            var templateBytes = ms.ToArray();
            using var templateStream = new MemoryStream(templateBytes);
            using var template = new WordTemplate(templateStream);

            var data = new Newtonsoft.Json.Linq.JObject
            {
                ["InvoiceSerial"] = "AA/26E-00123",
                ["CustomerName"] = "Nguyen Van Binh",
                ["SignerName"] = "Tran Van Giam Doc"
            };

            using var output = template.Export(data);
            Assert.NotNull(output);

            using var outDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(output!, false);
            // Verify Header filled
            var outHeader = outDoc.MainDocumentPart!.HeaderParts.First().Header;
            Assert.Contains("AA/26E-00123", outHeader.InnerText);

            // Verify Body filled
            var outBody = outDoc.MainDocumentPart!.Document.Body!;
            Assert.Contains("Nguyen Van Binh", outBody.InnerText);
            Assert.Contains("Họ tên khách hàng: ................................................................", outBody.InnerText);

            // Verify Footer filled
            var outFooter = outDoc.MainDocumentPart!.FooterParts.First().Footer;
            Assert.Contains("TRAN VAN GIAM DOC", outFooter.InnerText);
        }
    }
}
