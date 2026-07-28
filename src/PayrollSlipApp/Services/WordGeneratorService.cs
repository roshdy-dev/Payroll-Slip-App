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
            var local = Path.Combine(exeDir, "Format Word.docx");
            if (File.Exists(local)) return local;

            // Fall back to the original project path
            var project = @"D:\Payroll Slip App\Format Word.docx";
            if (File.Exists(project)) return project;

            throw new FileNotFoundException(
                "Word template 'Format Word.docx' not found. " +
                "Place it next to the executable or at the original project path.");
        }
    }

    /// <summary>
    /// Normalized keys for fields that should display as plain integers
    /// (no thousand separator, no decimal places).
    /// </summary>
    private static readonly HashSet<string> IntegerFields = new()
    {
        "emphrcode", "oracode", "cid", "accountno", "branchcode"
    };

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
            var templateTable = body.Elements<Table>().First();
            var originalTableXml = templateTable.OuterXml;

            for (int i = 0; i < dept.Employees.Count; i++)
            {
                var emp = dept.Employees[i];

                if (i == 0)
                {
                    FillTable(templateTable, emp.RawData, emp.PayPeriod);
                    var tp = body.Elements<Paragraph>().First(p => p.InnerText.Contains("تفاصيل الراتب"));
                    // Replace old-style "يوليو 2026" and new-style <<Month>> <<Year>>
                    foreach (var t in tp.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                    {
                        t.Text = t.Text.Replace("يوليو 2026", emp.PayPeriod);
                        t.Text = t.Text.Replace("<<Month>>", arabicMonth(emp.PayPeriod));
                        t.Text = t.Text.Replace("<<Year>>", arabicYear(emp.PayPeriod));
                    }
                }
                else
                {
                    body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                    body.Append(new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        new Run(new RunProperties(
                            new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                            new Bold(), new FontSize { Val = "44" }
                        ), new Text($"تفاصيل الراتب الشهري عن شهر {emp.PayPeriod}")
                            { Space = SpaceProcessingModeValues.Preserve })));
                    var tc = new Table(originalTableXml);
                    FillTable(tc, emp.RawData, emp.PayPeriod);
                    body.Append(tc);
                }
            }

            doc.MainDocumentPart.Document.Save();
            files.Add(path);
        }
        return files;
    }

    private static void FillTable(Table table, Dictionary<string, string> rawData, string payPeriod)
    {
        // Split "يوليو 2026" → month="يوليو", year="2026"
        var parts = payPeriod.Split(' ');
        var arabicMonth = parts.Length > 0 ? parts[0] : "";
        var year = parts.Length > 1 ? parts[1] : "";

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

                // Also handle combined <<Month>> or «Month» by stripping all wrappers
                var txt = te.Text.Trim('<', '>', '\u00AB', '\u00BB').Trim();

                // Handle Month and Year placeholders
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

                // Normalize the field name and look up in RawData
                var key = ExcelReaderService.NormalizeKey(te.Text);
                if (rawData.TryGetValue(key, out var value))
                {
                    if (IntegerFields.Contains(key))
                    {
                        // Integer fields: strip decimals, no thousand separators
                        if (decimal.TryParse(value, System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out var id))
                            value = ((long)id).ToString();
                    }
                    else if (decimal.TryParse(value, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var d))
                    {
                        value = d.ToString("#,##0.00");
                    }
                    te.Text = string.IsNullOrWhiteSpace(value) ? "—" : value;
                    te.Space = SpaceProcessingModeValues.Preserve;
                }
            }
        }
    }

    private static string Safe(string n)
    {
        var inv = Path.GetInvalidFileNameChars();
        var s = new string(n.Where(c => !inv.Contains(c)).ToArray());
        return string.IsNullOrWhiteSpace(s) ? "Department" : s;
    }

    private static string arabicMonth(string payPeriod) =>
        payPeriod.Split(' ').FirstOrDefault() ?? "";

    private static string arabicYear(string payPeriod) =>
        payPeriod.Split(' ').Skip(1).FirstOrDefault() ?? "";
}
