using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PayrollSlipApp.Models;

namespace PayrollSlipApp.Services;

/// <summary>
/// Generates Word payslips by cloning the template and replacing «FieldName»
/// placeholders using the RawData dictionary on each EmployeePayroll.
/// RawData is populated by ExcelReaderService — no Excel re-reading needed.
/// </summary>
public class WordGeneratorService
{
    /// <summary>
    /// Resolves the template path dynamically:
    /// 1. Next to the executable (for published self-contained apps)
    /// 2. Original project path (for development/debug)
    /// </summary>
    private static string TemplatePath
    {
        get
        {
            // Try next to the executable first
            var exeDir = AppDomain.CurrentDomain.BaseDirectory;
            var local = Path.Combine(exeDir, "Word_Format.docx");
            if (File.Exists(local)) return local;

            // Fall back to the original project path
            var project = @"D:\Payroll Slip App\Word_Format.docx";
            if (File.Exists(project)) return project;

            throw new FileNotFoundException(
                "Word template 'Word_Format.docx' not found. " +
                "Place it next to the executable or at the original project path.");
        }
    }

    /// <summary>
    /// Normalized column-name → FormatDataType, built from AppConfig.FormatDataTypes.
    /// </summary>
    private readonly Dictionary<string, FormatDataType> _formatTypeMap;

    /// <summary>Arabic month names (index 0 = January).</summary>
    private static readonly string[] ArabicMonths =
        { "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
          "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر" };

    /// <summary>English month names (index 0 = January).</summary>
    private static readonly string[] EnglishMonths =
        { "January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December" };

    public WordGeneratorService(AppConfig config)
    {
        _formatTypeMap = BuildFormatTypeMap(config);
    }

    /// <summary>
    /// Builds a normalized lookup from FormatDataTypes for O(1) formatting decisions.
    /// </summary>
    private static Dictionary<string, FormatDataType> BuildFormatTypeMap(AppConfig config)
    {
        var map = new Dictionary<string, FormatDataType>();
        if (config.FormatDataTypes == null) return map;

        foreach (var fdt in config.FormatDataTypes)
        {
            if (!string.IsNullOrWhiteSpace(fdt.ColumnName))
            {
                var key = ExcelReaderService.NormalizeKey(fdt.ColumnName);
                map[key] = fdt;
            }
        }
        return map;
    }

    public List<string> GenerateWordDocuments(List<DepartmentGroup> groups, string outputDir)
    {
        Directory.CreateDirectory(outputDir);
        var files = new List<string>();

        foreach (var dept in groups)
        {
            var path = Path.Combine(outputDir, $"Payslip_{Safe(dept.DepartmentName)}.docx");
            File.Copy(TemplatePath, path, true);

            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document.Body!;

            // Validate the template has a table
            var tables = body.Elements<Table>().ToList();
            if (tables.Count == 0)
                throw new InvalidOperationException(
                    "The Word template does not contain any table. " +
                    "The template must have at least one table with «FieldName» placeholders.");

            var templateTable = tables[0];
            var originalTableXml = templateTable.OuterXml;

            for (int i = 0; i < dept.Employees.Count; i++)
            {
                var emp = dept.Employees[i];

                if (i == 0)
                {
                    // First employee: fill the template's table in-place (preserves all formatting)
                    FillTable(templateTable, emp.RawData, emp.PayPeriod);
                }
                else
                {
                    // Additional employees: page break + clone the original table
                    // The cloned table includes the title row with correct formatting —
                    // no separate title paragraph needed (avoids formatting mismatches)
                    body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                    var clonedTable = new Table(originalTableXml);
                    FillTable(clonedTable, emp.RawData, emp.PayPeriod);
                    body.Append(clonedTable);
                }
            }

            // ── Clean up: remove empty paragraphs from the body ──
            // The template may have trailing empty paragraphs that cause blank pages.
            // Page-break paragraphs are preserved.
            var emptyParas = body.Elements<Paragraph>()
                .Where(p => string.IsNullOrWhiteSpace(p.InnerText)
                         && !p.Descendants<Break>().Any(b => b.Type == BreakValues.Page))
                .ToList();
            foreach (var ep in emptyParas)
                ep.Remove();

            doc.MainDocumentPart.Document.Save();
            files.Add(path);
        }
        return files;
    }

    private void FillTable(Table table, Dictionary<string, string> rawData, string payPeriod)
    {
        // Split "يوليو 2026" → month="يوليو", year="2026"
        var parts = payPeriod.Split(' ');
        var arabicMonth = parts.Length > 0 ? parts[0] : "";
        var year = parts.Length > 1 ? parts[1] : "";

        // Determine the month index (0-based) from the Arabic month name
        int monthIdx = Array.IndexOf(ArabicMonths, arabicMonth);
        var englishMonth = monthIdx >= 0 ? EnglishMonths[monthIdx] : arabicMonth;

        foreach (var para in table.Descendants<Paragraph>().ToList())
        {
            para.Elements<FieldCode>().ToList().ForEach(f => f.Remove());
            para.Elements<FieldChar>().ToList().ForEach(f => f.Remove());

            foreach (var run in para.Elements<Run>().ToList())
            {
                var te = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>();
                if (te == null) continue;

                // Clear « » and << >> wrapper chars (guillemets OR angle brackets)
                if (te.Text == "\u00AB" || te.Text == "\u00BB" ||
                    te.Text == "<" || te.Text == ">")
                {
                    te.Text = "";
                    continue;
                }

                // Strip wrappers to get the bare placeholder name
                var txt = te.Text.Trim('<', '>', '\u00AB', '\u00BB').Trim();
                var key = ExcelReaderService.NormalizeKey(te.Text);

                // ── Check config for PayPeriod_Month / PayPeriod_Year ──
                if (_formatTypeMap.TryGetValue(key, out var fdt))
                {
                    var dt = (fdt.DataType ?? "").ToLowerInvariant();

                    if (dt == "payperiod_month")
                    {
                        var lang = (fdt.Language ?? "").Trim();
                        te.Text = lang.Equals("English", StringComparison.OrdinalIgnoreCase)
                            ? englishMonth : arabicMonth;
                        te.Space = SpaceProcessingModeValues.Preserve;
                        continue;
                    }

                    if (dt == "payperiod_year")
                    {
                        te.Text = year;
                        te.Space = SpaceProcessingModeValues.Preserve;
                        continue;
                    }
                }

                // ── Fallback: handle Month/Year without config (backward compat) ──
                if (txt.Equals("Month", StringComparison.OrdinalIgnoreCase))
                {
                    te.Text = arabicMonth;
                    te.Space = SpaceProcessingModeValues.Preserve;
                    continue;
                }
                if (txt.Equals("Year", StringComparison.OrdinalIgnoreCase))
                {
                    te.Text = year;
                    te.Space = SpaceProcessingModeValues.Preserve;
                    continue;
                }

                // ── Normal RawData lookup with config-based formatting ──
                if (rawData.TryGetValue(key, out var value))
                {
                    var dataType = fdt?.DataType?.ToLowerInvariant();
                    value = FormatValue(value, dataType);
                    te.Text = string.IsNullOrWhiteSpace(value) ? "—" : value;
                    te.Space = SpaceProcessingModeValues.Preserve;
                }
            }
        }
    }

    /// <summary>
    /// Formats a raw value according to the configured data type.
    /// Falls back to auto-detection: decimal → "#,##0.00", otherwise as-is.
    /// </summary>
    private static string FormatValue(string rawValue, string? dataType)
    {
        if (string.IsNullOrWhiteSpace(rawValue))
            return "—";

        switch (dataType)
        {
            case "integer":
                if (decimal.TryParse(rawValue, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var intVal))
                    return ((long)intVal).ToString();
                break;

            case "decimal":
                if (decimal.TryParse(rawValue, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var decVal))
                    return decVal.ToString("#,##0.00");
                break;

            case "date":
                if (DateTime.TryParse(rawValue, out var dateVal))
                    return dateVal.ToString("dd/MM/yyyy");
                // Also try parsing as Excel serial date number
                if (double.TryParse(rawValue, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var serial)
                    && serial > 1 && serial < 300000)
                {
                    var excelEpoch = new DateTime(1899, 12, 30);
                    return excelEpoch.AddDays(serial).ToString("dd/MM/yyyy");
                }
                break;

            case "datetime":
                if (DateTime.TryParse(rawValue, out var dtVal))
                    return dtVal.ToString("dd/MM/yyyy HH:mm");
                if (double.TryParse(rawValue, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var dtSerial)
                    && dtSerial > 1 && dtSerial < 300000)
                {
                    var excelEpoch = new DateTime(1899, 12, 30);
                    return excelEpoch.AddDays(dtSerial).ToString("dd/MM/yyyy HH:mm");
                }
                break;

            case "string":
                return rawValue;

            default:
                // No explicit type configured — auto-detect
                if (decimal.TryParse(rawValue, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var autoDec))
                    return autoDec.ToString("#,##0.00");
                break;
        }

        return rawValue;
    }

    private static string Safe(string n)
    {
        var inv = Path.GetInvalidFileNameChars();
        var s = new string(n.Where(c => !inv.Contains(c)).ToArray());
        return string.IsNullOrWhiteSpace(s) ? "Department" : s;
    }
}
