using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordTemplater;

namespace WordTemplater.Desktop;

public partial class MainWindow : Window
{
    private string _lastExportedFilePath = string.Empty;

    private const string SampleLogoBase64 =
        "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";

    public MainWindow()
    {
        InitializeComponent();
        InitializeDefaultPaths();
        if (ComboTemplates != null)
        {
            ComboTemplates.SelectedIndex = 0;
        }
        LoadSampleDataForSelectedTemplate();
        AppendLog("WordTemplater Studio initialized successfully.");
    }

    private void InitializeDefaultPaths()
    {
        string defaultOutputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Output");
        if (!Directory.Exists(defaultOutputDir))
        {
            Directory.CreateDirectory(defaultOutputDir);
        }

        if (TxtOutputPath != null)
        {
            TxtOutputPath.Text = Path.Combine(defaultOutputDir, "Exported_Document.docx");
        }
    }

    private void OnTemplateSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (ComboTemplates == null || ComboTemplates.SelectedItem is not ComboBoxItem selectedItem) return;

        string tag = selectedItem.Tag?.ToString() ?? "showcase";
        if (PanelCustomTemplate != null) PanelCustomTemplate.IsVisible = (tag == "custom");

        switch (tag)
        {
            case "showcase":
                if (TxtTemplateInfo != null) TxtTemplateInfo.Text = "Comprehensive showcase: QR Code, 1D Barcode, Raster Image (Base64), Mermaid diagram, Rich HTML, Conditional branches (inline/comparison), Formatting & Repeating Table.";
                if (TxtWatermarkText != null) TxtWatermarkText.Text = "OFFICIAL";
                if (TxtWatermarkColor != null) TxtWatermarkColor.Text = "2563EB";
                break;
            case "preprinted":
                if (TxtTemplateInfo != null) TxtTemplateInfo.Text = "Pre-printed Bill / Voucher: Uses transparent Text Shapes precisely positioned over dotted lines (...) in Header, Body, and Footer to guarantee zero displacement.";
                if (TxtWatermarkText != null) TxtWatermarkText.Text = "ORIGINAL BILL";
                if (TxtWatermarkColor != null) TxtWatermarkColor.Text = "94A3B8";
                break;
            case "invoice":
                if (TxtTemplateInfo != null) TxtTemplateInfo.Text = "Template: Invoice with table cloning, zebra striping, currency & loop metadata (_index, _first, _last).";
                if (TxtWatermarkText != null) TxtWatermarkText.Text = "PAID";
                if (TxtWatermarkColor != null) TxtWatermarkColor.Text = "10B981";
                break;
            case "employee":
                if (TxtTemplateInfo != null) TxtTemplateInfo.Text = "Template: Employee Profile testing deep JSON property paths (Employee.Address.City) and evaluators.";
                if (TxtWatermarkText != null) TxtWatermarkText.Text = "CONFIDENTIAL";
                if (TxtWatermarkColor != null) TxtWatermarkColor.Text = "DC2626";
                break;
            case "custom":
                if (TxtTemplateInfo != null) TxtTemplateInfo.Text = "Custom Template: Choose any .docx file with standard MergeFields or «field» syntax.";
                break;
        }

        LoadSampleDataForSelectedTemplate();
    }

    private void LoadSampleDataForSelectedTemplate()
    {
        if (ComboTemplates == null || TxtJsonData == null || ComboTemplates.SelectedItem is not ComboBoxItem selectedItem) return;

        string tag = selectedItem.Tag?.ToString() ?? "showcase";

        if (tag == "showcase")
        {
            var showcaseData = new
            {
                DocId = 4829,
                GeneratedAt = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                IsVerified = true,
                HasBonus = true,
                Score = 95,
                ApprovalStatus = true,
                BarcodePayload = "WT-DOCX-998822",
                QrPayload = "https://github.com/luuducly/WordTemplater",
                LogoImage = SampleLogoBase64,
                ArchitectureDiagram = "graph TD\n  A[Client Input] --> B(WordTemplater Core)\n  B --> C{Template Parser}\n  C -->|Tables| D[Table Cloner]\n  C -->|QR / Barcode| E[SkiaSharp Generator]\n  C -->|Mermaid| F[Kroki Renderer]\n  D & E & F --> G[OpenXML .docx Package]",
                HtmlNotes = "<h3>Enterprise Certification Summary</h3><p>This document was compiled with <b>WordTemplater v2.1.6</b>. Key capabilities:</p><ul><li><b>Native OpenXML</b>: Zero Microsoft Word/Office dependency</li><li><b>Rich Media</b>: Direct generation of QR Codes, Barcodes, Mermaid diagrams &amp; Images</li><li><b>Cross-Platform</b>: Runs identically on macOS, Linux, and Windows</li></ul>",
                Products = new object[]
                {
                    new { Sku = "SRV-K8S-01", Name = "Managed Kubernetes Cluster Pool (16 vCPU)", Category = "Infrastructure", Qty = 3, Price = 450.00, LineTotal = 1350.00 },
                    new { Sku = "STG-NVME-02", Name = "High-Speed NVMe Storage Block (4TB)", Category = "Storage", Qty = 2, Price = 220.00, LineTotal = 440.00 },
                    new { Sku = "SEC-WAF-03", Name = "Web Application Firewall & Shield", Category = "Security", Qty = 1, Price = 380.00, LineTotal = 380.00 },
                    new { Sku = "AI-INFER-04", Name = "GPU Neural Inference Engine (A100)", Category = "Compute", Qty = 1, Price = 1250.00, LineTotal = 1250.00 }
                },
                Subtotal = 3420.00,
                TaxRate = 0.10,
                Total = 3762.00,
                RefCode = "WT-2026"
            };

            TxtJsonData.Text = JsonConvert.SerializeObject(showcaseData, Formatting.Indented);
        }
        else if (tag == "preprinted")
        {
            var preprintedData = new
            {
                // 1. Header Fields (Text Shapes over dots in Header)
                InvoiceForm = "01GTKT0/001",
                InvoiceSerial = "AA/26E",
                InvoiceNo = 4589,
                BillDay = DateTime.Now.ToString("dd"),
                BillMonth = DateTime.Now.ToString("MM"),

                // 2. Customer / Buyer Body Fields (Text Shapes over dots in Body)
                CustomerName = "Ông Nguyễn Hoàng Nam",
                CustomerCompany = "Công ty Cổ phần Công nghệ & Viễn thông Quốc tế",
                CustomerTaxCode = "0109887766-001",
                CustomerAddress = "Tòa nhà Center Point, 106 Nguyễn Văn Trỗi, P. 8, Phú Nhuận, TP.HCM",
                PaymentMethod = "Chuyển khoản (TM/CK)",
                BankAccount = "19038899887701 (Techcombank)",

                // 3. Pre-printed Table Row Items
                Item1_Name = "Máy chủ Dell PowerEdge R750 (2x Intel Xeon Gold, 128GB RAM)",
                Item1_Unit = "Bộ",
                Item1_Qty = 2,
                Item1_Price = 85000000,
                Item1_Amount = 170000000,

                Item2_Name = "Thiết bị mạng Cisco Catalyst 9200L 48 Ports PoE+",
                Item2_Unit = "Chiếc",
                Item2_Qty = 1,
                Item2_Price = 42000000,
                Item2_Amount = 42000000,

                Item3_Name = "Gói dịch vụ triển khai hạ tầng & cấu hình bảo mật 24/7",
                Item3_Unit = "Gói",
                Item3_Qty = 1,
                Item3_Price = 18000000,
                Item3_Amount = 18000000,

                // 4. Financial Totals (Text Shapes over dots in Summary)
                Subtotal = 230000000,
                VatAmount = 23000000,
                TotalAmount = 253000000,
                TotalInWords = "Hai trăm năm mươi ba triệu đồng chẵn.",

                // 5. Footer Signatures (Text Shapes over dots in Footer)
                BuyerSign = "Nguyễn Hoàng Nam",
                DelivererSign = "Lê Thanh Tùng",
                ManagerSign = "Trần Đình Trọng"
            };

            TxtJsonData.Text = JsonConvert.SerializeObject(preprintedData, Formatting.Indented);
        }
        else if (tag == "invoice")
        {
            var invoiceData = new
            {
                InvoiceNumber = "INV-2026-8899",
                Date = DateTime.Now.ToString("yyyy-MM-dd"),
                Customer = new
                {
                    Name = "Global Logistics Ltd.",
                    Address = new { City = "San Francisco", Country = "United States" }
                },
                Items = new object[]
                {
                    new { Description = "Dedicated Cloud Server (16 vCPU, 64GB RAM)", Qty = 2, Price = 650.00, Tax = 0.10 },
                    new { Description = "Enterprise SSD Storage Volume (2TB)", Qty = 4, Price = 120.00, Tax = 0.10 },
                    new { Description = "Managed Kubernetes Cluster Support", Qty = 1, Price = 450.00, Tax = 0.08 },
                    new { Description = "Wildcard SSL Certificate (2 Years)", Qty = 1, Price = 180.00, Tax = 0.00 },
                    new { Description = "Automated Disaster Recovery Backup", Qty = 12, Price = 35.00, Tax = 0.08 }
                },
                Total = 2520.00,
                Note = "Thank you for your business! Payment is due within 30 days."
            };

            TxtJsonData.Text = JsonConvert.SerializeObject(invoiceData, Formatting.Indented);
        }
        else if (tag == "employee")
        {
            var employeeData = new
            {
                Title = "Annual Employee Performance Review",
                Date = DateTime.Now.ToString("yyyy-MM-dd"),
                Employee = new
                {
                    Code = "EMP-0078",
                    Name = "Nguyen Van An",
                    Department = "Cloud Infrastructure & Security",
                    Title = "Principal Architect",
                    Salary = 150000000,
                    JoinDate = "2021-03-15",
                    Address = new
                    {
                        Street = "123 Tech Park Blvd",
                        City = "Da Nang",
                        Country = "Vietnam"
                    }
                },
                Skills = new object[]
                {
                    new { Name = "C# / .NET 10", Level = "Expert", Years = 8 },
                    new { Name = "Kubernetes & Cloud Architecture", Level = "Advanced", Years = 6 },
                    new { Name = "High-Performance Systems & OpenXML", Level = "Expert", Years = 5 }
                },
                IsEligibleForBonus = true,
                Summary = "Outstanding technical execution and cross-platform architecture leadership."
            };

            TxtJsonData.Text = JsonConvert.SerializeObject(employeeData, Formatting.Indented);
        }
    }

    private void OnFormatJsonClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            string raw = TxtJsonData.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw)) return;

            var obj = JToken.Parse(raw);
            TxtJsonData.Text = obj.ToString(Formatting.Indented);
            AppendLog("JSON successfully formatted.");
        }
        catch (Exception ex)
        {
            AppendLog($"JSON format error: {ex.Message}");
        }
    }

    private void OnResetSampleDataClick(object? sender, RoutedEventArgs e)
    {
        LoadSampleDataForSelectedTemplate();
        AppendLog("Sample JSON reset to default.");
    }

    private async void OnBrowseCustomTemplateClick(object? sender, RoutedEventArgs e)
    {
        var topLevel = GetTopLevel(this);
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select Word Template (.docx)",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Word Document (*.docx)")
                {
                    Patterns = new[] { "*.docx" }
                }
            }
        });

        if (files.Count > 0)
        {
            TxtCustomTemplatePath.Text = files[0].Path.LocalPath;
            AppendLog($"Selected custom template: {files[0].Path.LocalPath}");
        }
    }

    private async void OnBrowseOutputPathClick(object? sender, RoutedEventArgs e)
    {
        var topLevel = GetTopLevel(this);
        if (topLevel == null) return;

        var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save Exported Document As",
            DefaultExtension = "docx",
            SuggestedFileName = "Filled_Document.docx",
            FileTypeChoices = new[]
            {
                new FilePickerFileType("Word Document (*.docx)")
                {
                    Patterns = new[] { "*.docx" }
                }
            }
        });

        if (file != null)
        {
            TxtOutputPath.Text = file.Path.LocalPath;
        }
    }

    private void OnExportDocumentClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            TxtStatusBadge.Text = "Processing...";
            TxtStatusBadge.Foreground = Avalonia.Media.Brushes.Orange;

            // 1. Validate JSON
            string jsonText = TxtJsonData.Text ?? "{}";
            JObject dataObj;
            try
            {
                dataObj = JObject.Parse(jsonText);
            }
            catch (Exception ex)
            {
                AppendLog($"[Error] Invalid JSON syntax: {ex.Message}");
                TxtStatusBadge.Text = "JSON Error";
                TxtStatusBadge.Foreground = Avalonia.Media.Brushes.Red;
                return;
            }

            // 2. Resolve Template Bytes
            byte[] templateBytes = GetTemplateBytes();
            if (templateBytes == null || templateBytes.Length == 0)
            {
                AppendLog("[Error] Failed to load template bytes.");
                TxtStatusBadge.Text = "Template Error";
                TxtStatusBadge.Foreground = Avalonia.Media.Brushes.Red;
                return;
            }

            // 3. Configure Export Options
            var options = new ExportOptions();
            if (ChkWatermark.IsChecked == true && !string.IsNullOrWhiteSpace(TxtWatermarkText.Text))
            {
                options.WatermarkText = TxtWatermarkText.Text.Trim();
                options.WatermarkColor = TxtWatermarkColor.Text?.Trim() ?? "C0C0C0";
                if (double.TryParse(TxtWatermarkAngle.Text, out double angle))
                {
                    options.WatermarkAngle = angle;
                }
            }

            if (ChkZebraStriping.IsChecked == true)
            {
                options.TableOddRowColor = TxtZebraOddColor.Text?.Trim() ?? "F1F5F9";
                options.TableEvenRowColor = TxtZebraEvenColor.Text?.Trim() ?? "FFFFFF";
            }

            // 4. Determine Output Path
            string outputPath = TxtOutputPath.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(outputPath))
            {
                outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Output", $"Export_{DateTime.Now:yyyyMMdd_HHmmss}.docx");
                TxtOutputPath.Text = outputPath;
            }

            string? dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            // 5. Render
            var sw = Stopwatch.StartNew();
            using var ms = new MemoryStream(templateBytes);
            using var template = new WordTemplate(ms);
            using var outStream = template.Export(dataObj, options);
            if (outStream == null)
            {
                throw new InvalidOperationException("Export returned null stream.");
            }

            using (var fs = File.Create(outputPath))
            {
                outStream.Position = 0;
                outStream.CopyTo(fs);
            }
            sw.Stop();

            var fileInfo = new FileInfo(outputPath);
            _lastExportedFilePath = outputPath;
            BtnOpenFile.IsEnabled = true;

            TxtStatusBadge.Text = "Success";
            TxtStatusBadge.Foreground = Avalonia.Media.Brushes.LimeGreen;

            AppendLog($"[Success] Exported successfully in {sw.ElapsedMilliseconds}ms ({fileInfo.Length / 1024.0:F1} KB)");
            AppendLog($"Saved at: {outputPath}");
        }
        catch (Exception ex)
        {
            TxtStatusBadge.Text = "Failed";
            TxtStatusBadge.Foreground = Avalonia.Media.Brushes.Red;
            AppendLog($"[Exception] {ex.GetType().Name}: {ex.Message}");
        }
    }

    private static SimpleField F(string code) =>
        new SimpleField(new Run(new Text("«" + code + "»"))) { Instruction = " MERGEFIELD  " + code + "  \\* MERGEFORMAT " };

    private static Run T(string text, bool bold = false, string? color = null, string? size = null)
    {
        var runProps = new RunProperties();
        if (bold) runProps.Append(new Bold());
        if (!string.IsNullOrEmpty(color)) runProps.Append(new Color { Val = color });
        if (!string.IsNullOrEmpty(size)) runProps.Append(new FontSize { Val = size });
        return new Run(runProps, new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static Paragraph P(params OpenXmlElement[] children) => new Paragraph(children);

    private static TableCell C(OpenXmlElement content, string? bgColor = null)
    {
        var cell = new TableCell(content);
        if (!string.IsNullOrEmpty(bgColor))
        {
            cell.Append(new TableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = bgColor }));
        }
        return cell;
    }

    private byte[] GetTemplateBytes()
    {
        string tag = (ComboTemplates.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "showcase";

        if (tag == "custom")
        {
            string path = TxtCustomTemplatePath.Text?.Trim() ?? string.Empty;
            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"Custom template file not found: {path}");
            }
            return File.ReadAllBytes(path);
        }

        if (tag == "showcase")
        {
            return BuildShowcaseTemplate();
        }

        if (tag == "preprinted")
        {
            return BuildPreprintedBillTemplate();
        }

        if (tag == "invoice")
        {
            return BuildInvoiceTemplate();
        }

        return BuildEmployeeTemplate();
    }

    private static Run CreateTextShape(string fieldCode, double leftPt, double topPt, double widthPt, double heightPt, string fontColor = "0F172A", string fontSizeHalfPt = "22", bool bold = false, bool italic = false)
    {
        var shape = new DocumentFormat.OpenXml.Vml.Shape
        {
            Id = "Shape_" + Guid.NewGuid().ToString("N"),
            Style = $"position:absolute;margin-left:{leftPt:F1}pt;margin-top:{topPt:F1}pt;width:{widthPt:F1}pt;height:{heightPt:F1}pt;z-index:251658240;mso-position-horizontal:absolute;mso-position-vertical:absolute",
            Filled = false,
            Stroked = false
        };

        var txbx = new DocumentFormat.OpenXml.Vml.TextBox
        {
            Style = "mso-fit-shape-to-text:t"
        };

        var runProps = new RunProperties();
        if (bold) runProps.Append(new Bold());
        if (italic) runProps.Append(new Italic());
        if (!string.IsNullOrEmpty(fontColor)) runProps.Append(new Color { Val = fontColor });
        if (!string.IsNullOrEmpty(fontSizeHalfPt)) runProps.Append(new FontSize { Val = fontSizeHalfPt });

        var field = F(fieldCode);
        var innerRun = field.Descendants<Run>().FirstOrDefault();
        if (innerRun != null)
        {
            innerRun.RunProperties = (RunProperties)runProps.CloneNode(true);
        }

        var p = new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
            field
        );

        var txbxContent = new TextBoxContent(p);
        txbx.Append(txbxContent);
        shape.Append(txbx);

        return new Run(new Picture(shape));
    }

    public static byte[] BuildPreprintedBillTemplate()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document();

            // 1. Header with exact text shapes over pre-printed dots
            var headerPart = main.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(
                P(T("CÔNG TY CỔ PHẦN THIẾT BỊ & CÔNG NGHỆ Á CHÂU", true, "1E293B", "24")),
                P(T("MST: 0101234567-001  |  ĐC: 456 Hoàng Hoa Thám, Ba Đình, Hà Nội  |  Hotline: 1900-6868", false, "64748B", "18")),
                P(
                    T("Mẫu số: ....................       Ký hiệu: ....................       Số: ....................", false, "475569", "20"),
                    CreateTextShape("InvoiceForm", leftPt: 40, topPt: -2, widthPt: 85, heightPt: 16, fontColor: "1E293B", bold: true),
                    CreateTextShape("InvoiceSerial", leftPt: 175, topPt: -2, widthPt: 75, heightPt: 16, fontColor: "1E293B", bold: true),
                    CreateTextShape("padleft(InvoiceNo, 7, '0')", leftPt: 295, topPt: -3, widthPt: 90, heightPt: 18, fontColor: "DC2626", fontSizeHalfPt: "26", bold: true)
                ),
                P(
                    T("Ngày lập hóa đơn: Ngày ..... tháng ..... năm 2026", false, "475569", "20"),
                    CreateTextShape("BillDay", leftPt: 110, topPt: -2, widthPt: 25, heightPt: 16, fontColor: "1E293B", bold: true),
                    CreateTextShape("BillMonth", leftPt: 170, topPt: -2, widthPt: 25, heightPt: 16, fontColor: "1E293B", bold: true)
                )
            );

            // 2. Footer with exact text shapes for signatures
            var footerPart = main.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(
                P(
                    T("Người mua hàng                                 Người giao hàng                                 Thủ trưởng đơn vị", true, "334155", "20")
                ),
                P(
                    T("(Ký, ghi rõ họ tên)                             (Ký, ghi rõ họ tên)                             (Ký, đóng dấu, họ tên)", false, "64748B", "18")
                ),
                P(
                    T("...................................             ...................................             ...................................", false, "94A3B8", "20"),
                    CreateTextShape("BuyerSign", leftPt: 10, topPt: 15, widthPt: 130, heightPt: 18, fontColor: "1E293B", bold: true),
                    CreateTextShape("DelivererSign", leftPt: 175, topPt: 15, widthPt: 130, heightPt: 18, fontColor: "1E293B", bold: true),
                    CreateTextShape("ManagerSign:upper()", leftPt: 345, topPt: 15, widthPt: 140, heightPt: 18, fontColor: "1E293B", bold: true)
                )
            );

            var sectionProps = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(headerPart) },
                new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footerPart) }
            );

            // Table Header on Pre-printed Bill
            var tableHeader = new TableRow(
                C(P(T("STT", true, "1E293B")), bgColor: "F1F5F9"),
                C(P(T("Tên hàng hóa, dịch vụ", true, "1E293B")), bgColor: "F1F5F9"),
                C(P(T("ĐVT", true, "1E293B")), bgColor: "F1F5F9"),
                C(P(T("Số lượng", true, "1E293B")), bgColor: "F1F5F9"),
                C(P(T("Đơn giá", true, "1E293B")), bgColor: "F1F5F9"),
                C(P(T("Thành tiền (VNĐ)", true, "1E293B")), bgColor: "F1F5F9")
            );

            var row1 = new TableRow(
                C(P(T("1"))),
                C(P(F("Item1_Name"))),
                C(P(F("Item1_Unit"))),
                C(P(F("Item1_Qty:number('#,##0')"))),
                C(P(F("Item1_Price:currency(vi-VN)"))),
                C(P(F("Item1_Amount:currency(vi-VN)")))
            );

            var row2 = new TableRow(
                C(P(T("2"))),
                C(P(F("Item2_Name"))),
                C(P(F("Item2_Unit"))),
                C(P(F("Item2_Qty:number('#,##0')"))),
                C(P(F("Item2_Price:currency(vi-VN)"))),
                C(P(F("Item2_Amount:currency(vi-VN)")))
            );

            var row3 = new TableRow(
                C(P(T("3"))),
                C(P(F("Item3_Name"))),
                C(P(F("Item3_Unit"))),
                C(P(F("Item3_Qty:number('#,##0')"))),
                C(P(F("Item3_Price:currency(vi-VN)"))),
                C(P(F("Item3_Amount:currency(vi-VN)")))
            );

            // 3. Body with dotted background lines & overlay Text Shapes
            var body = new Body(
                P(T("HÓA ĐƠN BÁN HÀNG / GIÁ TRỊ GIA TĂNG", true, "0F172A", "32")),
                P(T("(Bản in phôi sẵn - Khắc phục sai lệch vị trí bằng Text Shape)", false, "64748B", "18")),
                P(),
                // Pre-printed Dotted Field 1: Customer Name
                P(
                    T("Họ tên người mua hàng: ............................................................................................................", false, "475569", "20"),
                    CreateTextShape("CustomerName", leftPt: 130, topPt: -2, widthPt: 320, heightPt: 18, fontColor: "0F172A", bold: true)
                ),
                // Pre-printed Dotted Field 2: Company Name
                P(
                    T("Tên đơn vị (Doanh nghiệp): ....................................................................................................", false, "475569", "20"),
                    CreateTextShape("CustomerCompany", leftPt: 140, topPt: -2, widthPt: 320, heightPt: 18, fontColor: "0F172A", bold: false)
                ),
                // Pre-printed Dotted Field 3: Tax Code
                P(
                    T("Mã số thuế: ............................................................................................................................", false, "475569", "20"),
                    CreateTextShape("CustomerTaxCode", leftPt: 75, topPt: -2, widthPt: 180, heightPt: 18, fontColor: "0F172A", bold: true)
                ),
                // Pre-printed Dotted Field 4: Address
                P(
                    T("Địa chỉ trụ sở / giao hàng: ......................................................................................................", false, "475569", "20"),
                    CreateTextShape("CustomerAddress", leftPt: 135, topPt: -2, widthPt: 330, heightPt: 18, fontColor: "0F172A", bold: false)
                ),
                // Pre-printed Dotted Field 5: Payment method & bank
                P(
                    T("Hình thức thanh toán: ........................................   Số tài khoản: .................................................", false, "475569", "20"),
                    CreateTextShape("PaymentMethod", leftPt: 125, topPt: -2, widthPt: 100, heightPt: 18, fontColor: "0F172A"),
                    CreateTextShape("BankAccount", leftPt: 295, topPt: -2, widthPt: 170, heightPt: 18, fontColor: "0F172A")
                ),
                P(),
                // Items Table
                new Table(tableHeader, row1, row2, row3),
                P(),
                // Financial Totals over Pre-printed dots
                P(
                    T("Cộng tiền hàng: ............................................................................................................................", false, "475569", "20"),
                    CreateTextShape("Subtotal:currency(vi-VN)", leftPt: 105, topPt: -2, widthPt: 200, heightPt: 18, fontColor: "0F172A", bold: true)
                ),
                P(
                    T("Tiền thuế GTGT (10%): ...............................................................................................................", false, "475569", "20"),
                    CreateTextShape("VatAmount:currency(vi-VN)", leftPt: 140, topPt: -2, widthPt: 200, heightPt: 18, fontColor: "0F172A", bold: true)
                ),
                P(
                    T("Tổng cộng tiền thanh toán: .....................................................................................................", true, "0F172A", "22"),
                    CreateTextShape("TotalAmount:currency(vi-VN)", leftPt: 160, topPt: -3, widthPt: 220, heightPt: 20, fontColor: "DC2626", fontSizeHalfPt: "24", bold: true)
                ),
                P(
                    T("Số tiền viết bằng chữ: .................................................................................................................", false, "475569", "20"),
                    CreateTextShape("TotalInWords", leftPt: 130, topPt: -2, widthPt: 330, heightPt: 18, fontColor: "0F172A", fontSizeHalfPt: "20", bold: true, italic: true)
                ),
                sectionProps
            );

            main.Document.Body = body;
        }
        return ms.ToArray();
    }

    public static byte[] BuildShowcaseTemplate()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();

            var prodHeader = new TableRow(
                C(P(T("#", true, "1E293B")), bgColor: "E2E8F0"),
                C(P(T("SKU", true, "1E293B")), bgColor: "E2E8F0"),
                C(P(T("Product / Service Description", true, "1E293B")), bgColor: "E2E8F0"),
                C(P(T("Category", true, "1E293B")), bgColor: "E2E8F0"),
                C(P(T("Qty", true, "1E293B")), bgColor: "E2E8F0"),
                C(P(T("Unit Price", true, "1E293B")), bgColor: "E2E8F0"),
                C(P(T("Line Total", true, "1E293B")), bgColor: "E2E8F0")
            );

            var prodRow = new TableRow(
                C(P(F("_index"))),
                C(P(F("Sku:upper()"))),
                C(P(F("Name"))),
                C(P(F("Category:lower()"))),
                C(P(F("Qty:number(#,##0)"))),
                C(P(F("Price:currency(en-US)"))),
                C(P(F("LineTotal:currency(en-US)")))
            );

            var body = new Body(
                // 1. Header Banner & Title
                P(T("WORDTEMPLATER ENTERPRISE SHOWCASE", true, "0F172A", "36")),
                P(T("Comprehensive Multi-Data-Type Test Suite for OpenXML Document Automation", false, "64748B", "20")),
                P(),

                // 2. Raster Image (Logo)
                P(T("1. Embedded Raster Image (Base64 PNG):", true, "1E293B", "22")),
                P(F("image(LogoImage)")),
                P(),

                // 3. Document Metadata & String Formatting
                P(T("2. Document Metadata & String Formatting:", true, "1E293B", "22")),
                P(T("Document ID (Padded 8 digits): ", true), F("padleft(DocId, 8, '0')"),
                  T("   |   Reference: ", true), F("RefCode:padright(12, '*')")),
                P(T("Generated At: ", true), F("date(GeneratedAt, 'dddd, dd MMMM yyyy HH:mm', 'en-US')")),
                P(),

                // 4. Machine-Readable Identifiers (QR Code & Barcode)
                P(T("3. Visual Machine-Readable Codes (QR Code & Barcode):", true, "1E293B", "22")),
                P(T("1D Tracking Barcode (Code128):", true, "334155")),
                P(F("barcode(BarcodePayload)")),
                P(T("Barcode Value: ", true), F("BarcodePayload")),
                P(),
                P(T("2D Verification QR Code:", true, "334155")),
                P(F("qrcode(QrPayload)")),
                P(T("Encoded Target URL: ", true), F("QrPayload")),
                P(),

                // 5. Conditional Logic & Dynamic Branches
                P(T("4. Conditional Logic & Dynamic Branches:", true, "1E293B", "22")),
                P(T("Verification Status (Inline Boolean): ", true),
                  F("if(IsVerified, 'VERIFIED & CERTIFIED (Digital Signature Valid)', 'UNVERIFIED DRAFT')")),
                P(T("Performance Incentive (Inline Boolean): ", true),
                  F("if(HasBonus, 'APPROVED - Performance Bonus Granted', 'NOT ELIGIBLE')")),
                P(T("VIP Tier Assessment (Inline Comparison Score >= 80): ", true),
                  F("if(Score >= 80, 'PLATINUM VIP (Tier 1)', 'STANDARD MEMBER')")),
                P(),

                // 6. Dynamic Repeating Table with Evaluators & Zebra Striping
                P(T("5. Dynamic Repeating Table with Evaluators & Zebra Striping:", true, "1E293B", "22")),
                P(F("table(Products)")),
                new Table(prodHeader, prodRow),
                P(F("endtable")),
                P(),

                // Financial Summary
                P(T("Subtotal: ", true), F("Subtotal:currency(en-US)"),
                  T("   |   Tax Rate: ", true), F("TaxRate:percentage(2)"),
                  T("   |   Net Total: ", true, "0F172A", "24"), F("Total:currency(en-US)")),
                P(),

                // 7. Rich Formatted HTML Content
                P(T("6. Rich Formatted HTML Content (AltChunk Ingestion):", true, "1E293B", "22")),
                P(F("html(HtmlNotes)")),
                P(),

                // 8. Mermaid Architecture Diagram
                P(T("7. Mermaid Architecture Diagram (Automated Kroki Service):", true, "1E293B", "22")),
                P(F("mermaid(ArchitectureDiagram)"))
            );

            main.Document = new Document(body);
        }
        return ms.ToArray();
    }

    public static byte[] BuildInvoiceTemplate()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var header = new TableRow(
                C(P(T("No.", true, "1E293B")), bgColor: "F8FAFC"),
                C(P(T("Description", true, "1E293B")), bgColor: "F8FAFC"),
                C(P(T("Qty", true, "1E293B")), bgColor: "F8FAFC"),
                C(P(T("Unit Price", true, "1E293B")), bgColor: "F8FAFC")
            );
            var row = new TableRow(
                C(P(F("_index"))),
                C(P(F("Description:upper()"))),
                C(P(F("Qty:number(#,##0)"))),
                C(P(F("Price:currency(en-US)")))
            );

            var body = new Body(
                P(T("COMMERCIAL INVOICE", true, "1E293B", "36")),
                P(T("Invoice Number: ", true), F("InvoiceNumber"), T("  |  Date: ", true), F("Date")),
                P(T("Customer: ", true), F("Customer.Name"), T("  |  Location: ", true), F("Customer.Address.City"), T(", "), F("Customer.Address.Country")),
                P(),
                P(F("table(Items)")),
                new Table(header, row),
                P(F("endtable")),
                P(),
                P(T("Total Amount: ", true, "0F172A", "24"), F("Total:currency(en-US)")),
                P(T("Payment Note: ", true), F("Note"))
            );
            main.Document = new Document(body);
        }
        return ms.ToArray();
    }

    public static byte[] BuildEmployeeTemplate()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var header = new TableRow(
                C(P(T("#", true, "0F172A")), bgColor: "F1F5F9"),
                C(P(T("Skill", true, "0F172A")), bgColor: "F1F5F9"),
                C(P(T("Proficiency", true, "0F172A")), bgColor: "F1F5F9"),
                C(P(T("Experience (Yrs)", true, "0F172A")), bgColor: "F1F5F9")
            );
            var row = new TableRow(
                C(P(F("_index"))),
                C(P(F("Name"))),
                C(P(F("Level:upper()"))),
                C(P(F("Years:number(#,##0)")))
            );

            var body = new Body(
                P(T("EMPLOYEE PERFORMANCE EVALUATION", true, "0F172A", "32")),
                P(T("Evaluation Date: ", true), F("Date")),
                P(),
                P(T("Employee Code: ", true), F("Employee.Code"), T("  |  Full Name: ", true), F("Employee.Name:upper()")),
                P(T("Department: ", true), F("Employee.Department")),
                P(T("Role / Title: ", true), F("Employee.Title")),
                P(T("Monthly Salary: ", true), F("Employee.Salary:currency(vi-VN)")),
                P(T("Join Date: ", true), F("Employee.JoinDate")),
                P(T("Office Location: ", true), F("Employee.Address.Street"), T(", "), F("Employee.Address.City"), T(", "), F("Employee.Address.Country")),
                P(),
                P(T("Technical Competencies:", true, "0F172A", "22")),
                P(F("table(Skills)")),
                new Table(header, row),
                P(F("endtable")),
                P(),
                P(T("Performance Review: ", true), F("Summary")),
                P(T("Bonus Recommendation: ", true), F("if(IsEligibleForBonus, 'APPROVED - Performance Bonus Awarded', 'NOT ELIGIBLE')"))
            );
            main.Document = new Document(body);
        }
        return ms.ToArray();
    }

    private void OnOpenFileClick(object? sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_lastExportedFilePath) || !File.Exists(_lastExportedFilePath))
        {
            AppendLog("[Warning] Output file does not exist.");
            return;
        }

        try
        {
            OpenWithDefaultApplication(_lastExportedFilePath);
            AppendLog($"Opened document: {_lastExportedFilePath}");
        }
        catch (Exception ex)
        {
            AppendLog($"[Error] Could not open file: {ex.Message}");
        }
    }

    private void OnOpenOutputFolderClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            string folder = Path.GetDirectoryName(TxtOutputPath.Text) ?? AppDomain.CurrentDomain.BaseDirectory;
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            OpenWithDefaultApplication(folder);
            AppendLog($"Opened folder: {folder}");
        }
        catch (Exception ex)
        {
            AppendLog($"[Error] Could not open folder: {ex.Message}");
        }
    }

    private static void OpenWithDefaultApplication(string path)
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Process.Start(new ProcessStartInfo("cmd", $"/c start \"\" \"{path}\"") { CreateNoWindow = true });
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            Process.Start("open", path);
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            Process.Start("xdg-open", path);
        }
    }

    private void AppendLog(string message)
    {
        if (TxtLogs == null) return;
        string timestamp = DateTime.Now.ToString("HH:mm:ss");
        string line = $"[{timestamp}] {message}";
        if (string.IsNullOrEmpty(TxtLogs.Text))
        {
            TxtLogs.Text = line;
        }
        else
        {
            TxtLogs.Text += Environment.NewLine + line;
        }
    }
}
