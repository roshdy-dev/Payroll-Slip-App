# Payroll Slip Generator

A cross-platform desktop application that generates professional payslips (Word & PDF) from Excel payroll data. Built with .NET 8 and Avalonia UI — runs on **Windows, macOS, and Linux**.

## 📥 Download (Windows)

**[⬇️ Download PayrollSlipGenerator-v1.0.zip](https://github.com/roshdy-dev/Payroll-Slip-App/releases/latest)**

1. Download the ZIP file
2. Extract to any folder
3. Run `PayrollSlipGenerator.exe` — no installation needed

> **Note:** The ZIP includes `Format Word.docx` (the payslip template) — keep it next to the EXE. Microsoft Word must be installed for PDF generation.

## ✨ Features

- 📊 **Excel Input** — Reads payroll data from `.xlsx`/`.xls` files with flexible column mapping
- 📝 **Word Output** — Generates `.docx` files with formatted payslips (one per department)
- 📄 **PDF Output** — Identical layout in PDF format
- 🏢 **Department Grouping** — Each department gets its own file with all employees
- 🖥️ **Cross-Platform GUI** — Native-looking interface on Windows, macOS, and Linux
- 📦 **Self-Contained** — Publish as a single executable, no .NET installation required

## 🚀 Quick Start

### Prerequisites (for development)

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

### Build & Run (Development)

```bash
cd "src/PayrollSlipApp"
dotnet run
```

### Publish Self-Contained Executable (No .NET Required)

**Windows:**

```batch
publish-windows.bat
```

**Linux:**

```bash
chmod +x publish-linux.sh
./publish-linux.sh
```

**macOS:**

```bash
chmod +x publish-macos.sh
./publish-macos.sh
```

The published executable in `publish/` can be copied to any machine — **no .NET runtime needed**.

## � Excel File Format

The first row of the Excel file must contain column headers. The application automatically detects columns by their header names:

| Column Header Examples                     | Maps To                  |
| ------------------------------------------ | ------------------------ |
| `EmployeeID`, `Emp ID`, `Staff ID`, `Code` | Employee ID              |
| `Employee Name`, `Name`, `Full Name`       | Employee Name            |
| `Department`, `Dept`, `Division`           | Department               |
| `Basic Salary`, `Basic`                    | Basic Salary             |
| `Housing Allowance`, `Housing`             | Housing Allowance        |
| `Transportation`, `Transport`              | Transportation Allowance |
| `Tax`, `Income Tax`                        | Tax Deduction            |
| `Social Insurance`, `Insurance`            | Social Insurance         |
| `Job Title`, `Designation`, `Position`     | Job Title                |
| `Pay Period`, `Month`                      | Pay Period               |
| `Bank Account`, `Bank`                     | Bank Account             |
| `Joining Date`, `Hire Date`                | Joining Date             |

## 📁 Output Structure

```
YourExcelFile.xlsx              ← Input file
Payslips_20260728_143022/       ← Auto-created output folder
├── Payslip_Engineering.docx    ← All Engineering employees
├── Payslip_Engineering.pdf
├── Payslip_Sales.docx          ← All Sales employees
├── Payslip_Sales.pdf
├── Payslip_HR.docx
└── Payslip_HR.pdf
```

Each file contains all employees in that department, each on a separate page.

## 🏗️ Project Structure

```
Payroll Slip App/
├── PayrollSlipApp.sln
├── publish-windows.bat          # Publish for Windows
├── publish-linux.sh             # Publish for Linux
├── publish-macos.sh             # Publish for macOS
└── src/PayrollSlipApp/
    ├── Program.cs               # Entry point
    ├── App.axaml/.cs            # Avalonia application setup
    ├── Models/
    │   └── PayrollModels.cs     # Employee & Department data models
    ├── Services/
    │   ├── ExcelReaderService.cs   # Excel parsing with dynamic column mapping
    │   ├── WordGeneratorService.cs # .docx generation (OpenXml)
    │   └── PdfGeneratorService.cs  # PDF generation (QuestPDF)
    ├── ViewModels/
    │   ├── ViewModelBase.cs        # MVVM base class
    │   └── MainViewModel.cs        # Main UI logic
    └── Views/
        ├── MainWindow.axaml        # UI layout (XAML)
        └── MainWindow.axaml.cs     # Code-behind
```

## 🔧 Tech Stack

| Component       | Technology                                                          | License         |
| --------------- | ------------------------------------------------------------------- | --------------- |
| UI Framework    | [Avalonia UI](https://avaloniaui.net/)                              | MIT             |
| MVVM            | [CommunityToolkit.Mvvm](https://github.com/CommunityToolkit/dotnet) | MIT             |
| Excel Reading   | [ClosedXML](https://closedxml.github.io/ClosedXML/)                 | MIT             |
| Word Generation | [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK)    | MIT             |
| PDF Generation  | [QuestPDF](https://www.questpdf.com/)                               | MIT (Community) |

## 🎨 Payslip Layout

Each payslip page includes:

1. **Company Header** — "EMPLOYEE PAYSLIP" title
2. **Employee Information** — ID, Name, Department, Job Title, Pay Period, Bank Account
3. **Earnings Table** — Basic Salary, Housing, Transportation, Other Allowances, Total
4. **Deductions Table** — Income Tax, Social Insurance, Other Deductions, Total
5. **Net Salary** — Highlighted green summary box
6. **Footer** — Generation timestamp

The Word and PDF outputs share identical formatting (fonts, colors, borders, spacing).

## ⚙️ Customization

- **Column Mappings** — Edit `ColumnMappings` dictionary in `ExcelReaderService.cs` to match your Excel headers
- **Colors & Styling** — Adjust color constants at the top of `WordGeneratorService.cs` and `PdfGeneratorService.cs`
- **Currency Format** — Change `FormatCurrency()` in both services
- **Page Size** — Modify `PageSize` in `WordGeneratorService.AddDocumentSettings()` (Word) and `PageSizes` in `PdfGeneratorService` (PDF)
