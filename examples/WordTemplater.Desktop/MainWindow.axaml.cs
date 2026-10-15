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

        string tag = selectedItem.Tag?.ToString() ?? "invoice";
        if (PanelCustomTemplate != null) PanelCustomTemplate.IsVisible = (tag == "custom");

        switch (tag)
        {
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

        string tag = selectedItem.Tag?.ToString() ?? "invoice";

        if (tag == "invoice")
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
        string tag = (ComboTemplates.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "invoice";

        if (tag == "custom")
        {
            string path = TxtCustomTemplatePath.Text?.Trim() ?? string.Empty;
            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"Custom template file not found: {path}");
            }
            return File.ReadAllBytes(path);
        }

        if (tag == "invoice")
        {
            return BuildInvoiceTemplate();
        }

        return BuildEmployeeTemplate();
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
