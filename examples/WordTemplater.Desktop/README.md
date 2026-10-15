# WordTemplater Studio (Desktop GUI App)

A modern, cross-platform desktop GUI application built with **Avalonia UI** and **.NET 8 / 9 / 10** to interactively test and preview `WordTemplater` document generation.

Runs natively on **macOS**, **Linux**, and **Windows**.

---

## Features

- 📑 **Built-in & Custom Templates**:
  - **Commercial Invoice**: Repeating table rows, zebra striping, currency formatting, loop metadata (`_index`).
  - **Employee Profile**: Deep JSON path navigation (`Employee.Address.City`), date formatting, upper/lower evaluators.
  - **Custom DOCX Template**: Browse and load any Microsoft Word template from your machine.
- 📝 **Live JSON Data Editor**:
  - Monospace code editor for inspecting and modifying template data payloads.
  - One-click **Format JSON** and **Reset Sample Data**.
- 🎨 **Dynamic Export Options**:
  - Configure **Watermarks** (Text, Hex Color, Rotation Angle).
  - Configure **Table Zebra Striping** (Custom Odd/Even row hex colors).
- ⚡ **Instant Rendering & Verification**:
  - Measures execution time in milliseconds and verifies file size.
  - **Open Exported Document**: Automatically launches the generated `.docx` file in your default OS document viewer (Microsoft Word, LibreOffice, Pages, etc.).

---

## How to Run

### Command Line (.NET CLI)

From repository root:
```bash
dotnet run --project examples/WordTemplater.Desktop/WordTemplater.Desktop.csproj
```

Or navigate to the project directory:
```bash
cd examples/WordTemplater.Desktop
dotnet run
```

---

## Technology Stack

- **Framework**: Avalonia UI 12 (Skia rendering on Metal for macOS, X11/Wayland for Linux, DirectX/Win32 for Windows).
- **Core Library**: `WordTemplater` (DocumentFormat.OpenXml).
- **JSON Engine**: `Newtonsoft.Json`.
- **Target Runtimes**: Multi-targeted for `.NET 8.0`, `.NET 9.0`, and `.NET 10.0`.
