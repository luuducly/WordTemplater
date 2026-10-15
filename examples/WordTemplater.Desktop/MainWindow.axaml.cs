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
            string builtInPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "InvoiceTemplate.docx");
            if (File.Exists(builtInPath))
            {
                return File.ReadAllBytes(builtInPath);
            }

            // Programmatic fallback template
            return BuildInvoiceTemplateInMemory();
        }

        // Employee template
        return BuildEmployeeTemplateInMemory();
    }

    private byte[] BuildInvoiceTemplateInMemory()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body();

            // Heading
            var headingRunProps = new RunProperties(new Bold(), new FontSize { Val = "36" }, new Color { Val = "1E293B" });
            body.Append(new Paragraph(new Run(headingRunProps, new Text("COMMERCIAL INVOICE") { Space = SpaceProcessingModeValues.Preserve })));

            body.Append(new Paragraph(new Run(new Text("Invoice #: «InvoiceNumber»  |  Date: «date(Date, 'dd/MM/yyyy')»"))));
            body.Append(new Paragraph(new Run(new Text("Customer: «Customer.Name» (City: «Customer.Address.City», «Customer.Address.Country»)"))));
            body.Append(new Paragraph());

            // Repeating Table
            var table = new Table();
            var tblPr = new TableProperties(new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = "CBD5E1" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "CBD5E1" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "E2E8F0" }
            ));
            table.Append(tblPr);

            // Header row
            var headerRow = new TableRow(
                CreateCell("No.", true, "1E293B", "F8FAFC"),
                CreateCell("Description", true, "1E293B", "F8FAFC"),
                CreateCell("Qty", true, "1E293B", "F8FAFC"),
                CreateCell("Price", true, "1E293B", "F8FAFC")
            );
            table.Append(headerRow);

            // Template row
            var dataRow = new TableRow(
                CreateCell("«table(Items)»«_index»"),
                CreateCell("«Description»"),
                CreateCell("«Qty»"),
                CreateCell("«currency(Price, 'en-US')»«endtable»")
            );
            table.Append(dataRow);

            body.Append(table);
            body.Append(new Paragraph());

            var totalRunProps = new RunProperties(new Bold());
            body.Append(new Paragraph(new Run(totalRunProps, new Text("Total Amount: «currency(Total, 'en-US')»"))));
            body.Append(new Paragraph(new Run(new Text("Note: «Note»"))));

            main.Document = new Document(body);
        }
        return ms.ToArray();
    }

    private byte[] BuildEmployeeTemplateInMemory()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body();

            var titleProps = new RunProperties(new Bold(), new FontSize { Val = "32" }, new Color { Val = "0F172A" });
            body.Append(new Paragraph(new Run(titleProps, new Text("«Title»") { Space = SpaceProcessingModeValues.Preserve })));

            body.Append(new Paragraph(new Run(new Text("Date: «date(Date, 'MMMM dd, yyyy')»"))));
            body.Append(new Paragraph());
            body.Append(new Paragraph(new Run(new Text("Employee Code: «Employee.Code»  |  Name: «upper(Employee.Name)»"))));
            body.Append(new Paragraph(new Run(new Text("Department: «Employee.Department»"))));
            body.Append(new Paragraph(new Run(new Text("Role: «Employee.Title»"))));
            body.Append(new Paragraph(new Run(new Text("Monthly Salary: «currency(Employee.Salary, 'vi-VN')»"))));
            body.Append(new Paragraph(new Run(new Text("Join Date: «date(Employee.JoinDate, 'dd/MM/yyyy')»"))));
            body.Append(new Paragraph(new Run(new Text("Office: «Employee.Address.Street», «Employee.Address.City», «Employee.Address.Country»"))));
            body.Append(new Paragraph());

            var compProps = new RunProperties(new Bold());
            body.Append(new Paragraph(new Run(compProps, new Text("Key Competencies:"))));

            var table = new Table();
            table.Append(new TableProperties(new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = "94A3B8" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "94A3B8" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "CBD5E1" }
            )));

            table.Append(new TableRow(
                CreateCell("#", true, "0F172A", "F1F5F9"),
                CreateCell("Skill", true, "0F172A", "F1F5F9"),
                CreateCell("Proficiency", true, "0F172A", "F1F5F9"),
                CreateCell("Experience (Yrs)", true, "0F172A", "F1F5F9")
            ));

            table.Append(new TableRow(
                CreateCell("«table(Skills)»«_index»"),
                CreateCell("«Name»"),
                CreateCell("«Level»"),
                CreateCell("«Years»«endtable»")
            ));

            body.Append(table);
            body.Append(new Paragraph());
            body.Append(new Paragraph(new Run(new Text("Performance Summary: «Summary»"))));
            body.Append(new Paragraph(new Run(new Text("Bonus Eligibility: «if(IsEligibleForBonus, 'YES - Approved', 'NO')»"))));

            main.Document = new Document(body);
        }
        return ms.ToArray();
    }

    private static TableCell CreateCell(string text, bool bold = false, string color = "000000", string? bgColor = null)
    {
        var runProps = new RunProperties();
        if (bold) runProps.Append(new Bold());
        if (!string.IsNullOrEmpty(color)) runProps.Append(new Color { Val = color });

        var p = new Paragraph(new Run(runProps, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        var cell = new TableCell(p);

        if (!string.IsNullOrEmpty(bgColor))
        {
            cell.Append(new TableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = bgColor }));
        }

        return cell;
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
