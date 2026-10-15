# WordTemplater

[![NuGet Version](https://img.shields.io/nuget/v/WordTemplater.svg?style=flat-square&color=blue)](https://www.nuget.org/packages/WordTemplater)
[![NuGet Downloads](https://img.shields.io/nuget/dt/WordTemplater.svg?style=flat-square&color=green)](https://www.nuget.org/packages/WordTemplater)
[![Target Frameworks](https://img.shields.io/badge/.NET-8.0%20%7C%209.0%20%7C%2010.0-purple.svg?style=flat-square)](https://dotnet.microsoft.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://github.com/luuducly/WordTemplater/blob/main/LICENSE.txt)

**WordTemplater** is a high-performance, cross-platform .NET library that turns Microsoft Word (`.docx`) documents into dynamic report templates. Design your templates visually in Microsoft Word using standard **MergeFields** (or `«field»` tags), and populate them with complex JSON or C# objects in milliseconds.

No Microsoft Office installation or COM Interop required. Runs seamlessly on Windows, Linux, macOS, Docker containers, and cloud serverless environments.

---

## Key Features

- 🎨 **Visual Template Design**: Design invoices, contracts, and certificates directly in Microsoft Word or LibreOffice.
- ⚡ **Zero Office Dependency**: 100% managed code built on OpenXML.
- 🐧 **True Cross-Platform**: Fully tested on Linux, macOS, and Windows. Multi-targeted for **.NET 8 (LTS)**, **.NET 9**, and **.NET 10**.
- 🧭 **Deep JSON Navigation**: Access nested fields (`Customer.Address.City`) and array indices (`Items[0].Name`) case-insensitively.
- 📊 **Table & Loop Automation**: Effortless table row cloning, dynamic alternating zebra-striping, and loop metadata (`_index`, `_first`, `_last`, `_even`, `_odd`).
- 🖼️ **Multi-Format Media Insertion**:
  - Images directly from file paths, HTTPS URLs, `byte[]`, `Stream`, or Base64 strings.
  - Dynamic QR Codes and Code128 Barcodes.
  - Flowcharts and diagrams via [Mermaid](https://mermaid.js.org/).
  - Raw HTML and nested sub-DOCX documents.
- 🛡️ **Document Watermarking**: Inject customizable diagonal or horizontal text watermarks into all sections/headers dynamically.
- 🧩 **Extensible Plugin Engine**: Register custom evaluators and formatters with a single interface (`IEvaluator`).

---

## Installation

```bash
dotnet add package WordTemplater --version 2.16
```

### PackageReference
```xml
<PackageReference Include="WordTemplater" Version="2.16" />
```

---

## Quick Start

```csharp
using System.IO;
using WordTemplater;

var data = new
{
    InvoiceNumber = "INV-2026-001",
    CreatedAt = DateTime.UtcNow,
    Customer = new
    {
        Name = "Acme Corporation",
        Address = new { City = "San Francisco" }
    },
    Items = new[]
    {
        new { Description = "Cloud Hosting", Qty = 1, Price = 1200.00 },
        new { Description = "Domain Registration", Qty = 2, Price = 15.00 }
    },
    CompanyLogo = "https://example.com/logo.png"
};

using (var templateStream = File.OpenRead("Templates/InvoiceTemplate.docx"))
using (var wordTemplate = new WordTemplate(templateStream))
{
    wordTemplate.SetWatermark("PAID", color: "22C55E");
    wordTemplate.EnableTableZebraStriping(oddColor: "F8FAFC", evenColor: "FFFFFF");

    using (var exportedStream = wordTemplate.Export(data))
    using (var outputStream = File.Create("Output/Invoice.docx"))
    {
        exportedStream.CopyTo(outputStream);
    }
}
```

---

## Template Syntax & Evaluators Cheat Sheet

### 1. Text & String Formatting

| Placeholder Syntax | Description | Example Input | Example Output |
| :--- | :--- | :--- | :--- |
| `«upper(Name)»` | Converts text to uppercase | `"acme corp"` | `ACME CORP` |
| `«lower(Code)»` | Converts text to lowercase | `"PROD-XYZ"` | `prod-xyz` |
| `«trim(Note)»` | Trims leading/trailing whitespace | `"  hello  "` | `hello` |
| `«sub(Text, 0, 4)»` | Substring `(start, length)` | `"WORDTEMPLATER"` | `WORD` |
| `«left(Code, 3)»` | First `N` characters | `"INV-9988"` | `INV` |
| `«right(Code, 4)»` | Last `N` characters | `"INV-9988"` | `9988` |
| `«replace(Text, '-', '/')»` | Replaces occurrences of string | `"2026-10-15"` | `2026/10/15` |
| `«padleft(Number, 6, '0')»` | Left-pads string to target length | `42` | `000042` |
| `«padright(Code, 8, '_')»` | Right-pads string to target length | `"A1"` | `A1______` |

---

### 2. Numbers & Currency

| Placeholder Syntax | Description | Example Input | Example Output |
| :--- | :--- | :--- | :--- |
| `«currency(Amount)»` | Formats as currency using current/specified culture | `1500000` | `$1,500,000.00` |
| `«currency(Amount, 'vi-VN')»` | Currency with explicit culture | `1500000` | `1.500.000 ₫` |
| `«number(Score, '#,##0.00')»` | Formats number with custom format pattern | `12345.6` | `12,345.60` |
| `«number(Score, '#,##0', 'de-DE')»` | Number format with target culture | `1234567` | `1.234.567` |
| `«percentage(Ratio)»` | Formats number as percentage | `0.85` | `85.00%` |

---

### 3. Dates & Times

| Placeholder Syntax | Description | Example Input | Example Output |
| :--- | :--- | :--- | :--- |
| `«date(CreatedAt, 'dd/MM/yyyy')»` | Formats date with pattern | `2026-10-15T09:00:00` | `15/10/2026` |
| `«date(CreatedAt, 'yyyy-MM-dd HH:mm')»` | Formats date and time | DateTime object | `2026-10-15 09:00` |
| `«date(CreatedAt, 'MMMM yyyy', 'vi-VN')»` | Formats date with localized culture | `2026-10-15` | `tháng mười 2026` |

---

### 4. Logic & Fallbacks

| Placeholder Syntax | Description | Example Input | Example Output |
| :--- | :--- | :--- | :--- |
| `«default(Phone, 'N/A')»` | Returns fallback value if null, empty, or missing | `null` | `N/A` |
| `«coalesce(TaxCode, 'Unregistered')»` | Synonym for `default` | `""` | `Unregistered` |
| `«if(IsVIP, 'Gold Member', 'Standard')»` | Outputs different strings depending on boolean | `true` | `Gold Member` |

---

### 5. Deep JSON & Property Navigation

WordTemplater seamlessly resolves nested properties and array indices:

```text
«Customer.Company.LegalName»
«Customer.BillingAddress.City»
«Orders[0].TrackingCode»
«Department.Managers[0].Contact.Email»
```

*Property name matching is case-insensitive.*

---

### 6. Repeating Lists & Loops

```text
«loop(Employees)»
• Employee: «Name» — Role: «Role» (Row #«_index»)
«endloop»
```

#### Loop Metadata Variables
- `«_index»`: 1-based index of the current item (`1, 2, 3...`)
- `«_first»`: `true` if current item is the first element
- `«_last»`: `true` if current item is the last element
- `«_even»`: `true` if `_index` is even (`2, 4, 6...`)
- `«_odd»`: `true` if `_index` is odd (`1, 3, 5...`)

---

### 7. Repeating Tables & Zebra Striping

| No. | Description | Qty | Unit Price |
| :--- | :--- | :--- | :--- |
| `«table(Items)»«_index»` | `«Description»` | `«Qty»` | `«currency(Price)»«endtable»` |

Enable alternating shading in template:
`«table(Items, 'F8FAFC', 'FFFFFF')»`

Or via C# API:
`wordTemplate.EnableTableZebraStriping("F8FAFC", "FFFFFF");`

---

### 8. Images, Barcodes & Media

| Syntax | Supported Data Formats | Description |
| :--- | :--- | :--- |
| `«image(Avatar)»` | Local file path, HTTPS URL, `byte[]`, `Stream`, or Base64 string | Inserts image directly preserving aspect ratio |
| `«qrcode(Url)»` | String / URL | Automatically renders a QR code image |
| `«barcode(Sku)»` | String / Numeric code | Generates a 1D Code128 barcode |
| `«mermaid(Diagram)»` | Mermaid graph syntax text | Automatically renders Mermaid text into a diagram image |
| `«html(Content)»` | HTML string | Converts HTML into native OpenXML Word paragraphs |
| `«word(DocxData)»` | Base64-encoded `.docx` file | Merges an entire sub-document into current position |

---

### 9. Custom Evaluators

```csharp
using WordTemplater;

public class SlugifyEvaluator : IEvaluator
{
    public string? Evaluate(object? value, List<object> parameters)
    {
        if (value == null) return null;
        var text = value.ToString() ?? "";
        return text.ToLowerInvariant().Replace(" ", "-");
    }
}

// Register evaluator
wordTemplate.RegisterEvaluator("slugify", new SlugifyEvaluator());
```

---

## Project Repository & Documentation

For full source code, examples, and issues, visit:
[https://github.com/luuducly/WordTemplater](https://github.com/luuducly/WordTemplater)
