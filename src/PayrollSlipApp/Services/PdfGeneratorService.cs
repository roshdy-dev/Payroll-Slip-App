using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using PayrollSlipApp.Models;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace PayrollSlipApp.Services;

/// <summary>
/// Generates PDF payslips using parallel processing for maximum speed.
/// 
/// Strategy:
/// 1. Generate Word files in parallel per department (CPU-bound, file I/O)
/// 2. Convert Word→PDF in parallel using multiple Word COM instances
///    - Each thread owns its own Word.Application (STA-required)
///    - Concurrency capped at ~4 to avoid memory thrashing
/// </summary>
public class PdfGeneratorService
{
    /// <summary>Max concurrent Word COM instances. Higher = faster but more RAM.</summary>
    private const int MaxWordInstances = 4;

    private readonly AppConfig _config;

    public PdfGeneratorService(AppConfig config)
    {
        _config = config;
    }

    static PdfGeneratorService() { QuestPDF.Settings.License = LicenseType.Community; }

    public List<string> GeneratePdfDocuments(List<DepartmentGroup> groups, string outputDir)
    {
        Directory.CreateDirectory(outputDir);

        // ── Phase 1: Generate Word files in parallel ──
        var wordFiles = new ConcurrentBag<string>();
        Parallel.ForEach(groups, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, dept =>
        {
            var wordGen = new WordGeneratorService(_config);
            var files = wordGen.GenerateWordDocuments(new List<DepartmentGroup> { dept }, outputDir);
            foreach (var f in files) wordFiles.Add(f);
        });

        var wordPaths = wordFiles.ToList();

        // ── Phase 2: Convert Word → PDF in parallel ──
        var pdfPaths = new ConcurrentBag<string>();
        bool wordAvailable = IsWordAvailable();

        if (wordAvailable)
        {
            // Parallel Word COM conversion — each thread gets its own Word instance
            Parallel.ForEach(wordPaths, new ParallelOptions { MaxDegreeOfParallelism = MaxWordInstances }, wordPath =>
            {
                var pdfPath = Path.ChangeExtension(wordPath, ".pdf");
                ConvertWithWord(wordPath, pdfPath);
                pdfPaths.Add(pdfPath);
            });
        }
        else
        {
            // Parallel QuestPDF fallback
            Parallel.ForEach(groups, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, dept =>
            {
                var pdfPath = Path.Combine(outputDir, $"Payslip_{SafeFileName(dept.DepartmentName)}.pdf");
                GenerateQuestPdf(dept, pdfPath);
                pdfPaths.Add(pdfPath);
            });
        }

        return pdfPaths.ToList();
    }

    // ═══════════════════════════════════════════════
    //  WORD COM (pixel-perfect, thread-safe per instance)
    // ═══════════════════════════════════════════════

    private static bool IsWordAvailable()
    {
        try { return Type.GetTypeFromProgID("Word.Application") != null; }
        catch { return false; }
    }

    /// <summary>
    /// Converts a single .docx to .pdf using Word COM.
    /// Each call creates its own Word instance (safe for parallel use).
    /// </summary>
    private static void ConvertWithWord(string docxPath, string pdfPath)
    {
        object wordApp = null!, docs = null!, doc = null!;
        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application")
                ?? throw new InvalidOperationException("Word not available");

            wordApp = Activator.CreateInstance(wordType)!;
            wordType.InvokeMember("Visible", BindingFlags.SetProperty, null, wordApp, new object[] { false });
            // Disable alerts (e.g., "file is read-only" dialogs)
            wordType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, wordApp, new object[] { 0 }); // wdAlertsNone
            // Suppress background printing for speed
            wordType.InvokeMember("Options", BindingFlags.GetProperty, null, wordApp, null);

            var fullDocx = Path.GetFullPath(docxPath);
            docs = wordType.InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null)!;
            var docsType = docs.GetType();
            doc = docsType.InvokeMember("Open", BindingFlags.InvokeMethod, null, docs,
                new object[] { fullDocx, true, false })!;

            var docType = doc.GetType();
            var fullPdf = Path.GetFullPath(pdfPath);

            // SaveAs PDF format = 17
            docType.InvokeMember("SaveAs2", BindingFlags.InvokeMethod, null, doc,
                new object[] { fullPdf, 17 });

            docType.InvokeMember("Close", BindingFlags.InvokeMethod, null, doc, new object[] { false });
        }
        finally
        {
            if (doc != null) Marshal.ReleaseComObject(doc);
            if (docs != null) Marshal.ReleaseComObject(docs);
            if (wordApp != null)
            {
                wordApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, wordApp, new object[] { false });
                Marshal.ReleaseComObject(wordApp);
            }
        }
    }

    // ═══════════════════════════════════════════════
    //  QUESTPDF FALLBACK
    // ═══════════════════════════════════════════════

    private static void GenerateQuestPdf(DepartmentGroup dept, string pdfPath)
    {
        const float margin = 50.4f;
        Document.Create(container =>
        {
            foreach (var emp in dept.Employees)
            {
                var d = emp.RawData;
                string V(string k) => d.TryGetValue(ExcelReaderService.NormalizeKey(k), out var v) ? v : "—";
                decimal D(string k) => decimal.TryParse(V(k),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var x) ? x : 0;
                string F(decimal v) => v == 0 ? "—" : v.ToString("#,##0.00");
                string I(string k) { var r = V(k); return decimal.TryParse(r, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var n) ? ((long)n).ToString() : r; }

                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.MarginVertical(margin); page.MarginHorizontal(margin);
                    page.DefaultTextStyle(x => x.FontFamily("Arial").FontSize(8));
                    page.Header().Column(h =>
                    {
                        h.Item().Text("كود الشركة: 100000").FontSize(8);
                        h.Item().Text("كود الموظف: 100100").FontSize(8);
                        h.Item().PaddingBottom(6).Text("-175260-123825").FontSize(8);
                    });
                    page.Content().Column(col =>
                    {
                        col.Item().AlignCenter().Text($"تفاصيل الراتب الشهري عن شهر {emp.PayPeriod}").FontFamily("Arial").FontSize(22).Bold();
                        col.Item().Height(8);
                        col.Item().Table(tbl =>
                        {
                            tbl.ColumnsDefinition(c =>
                            {
                                c.ConstantColumn(109.9f); c.ConstantColumn(7.8f);
                                c.ConstantColumn(121.5f); c.ConstantColumn(174.5f);
                                c.ConstantColumn(138.7f); c.ConstantColumn(188.3f);
                            });
                            IContainer Cell(uint s, string? bg = null) => tbl.Cell().ColumnSpan(s).Border(0.5f).BorderColor(Colors.Black).Background(bg ?? Colors.White).PaddingVertical(3).PaddingHorizontal(4);
                            void Info(string l, string v) { Cell(1).AlignRight().Text(l).FontSize(10).Bold(); Cell(5).AlignRight().Text(v).FontSize(10); }
                            void RH() { Cell(2, "#FF0000").AlignCenter().Text("الاستحقاقات").FontSize(10).Bold().FontColor(Colors.White); Cell(2, "#FF0000").AlignCenter().Text("القيمة").FontSize(10).Bold().FontColor(Colors.White); Cell(1, "#FF0000").AlignCenter().Text("الاستقطاعات").FontSize(10).Bold().FontColor(Colors.White); Cell(1, "#FF0000").AlignCenter().Text("القيمة").FontSize(10).Bold().FontColor(Colors.White); }
                            void DR(string el, decimal ea, string dl, decimal da) { Cell(2).PaddingVertical(2).AlignRight().Text(el).FontSize(9); Cell(2).PaddingVertical(2).AlignRight().Text(ea != 0 ? F(ea) : "—").FontSize(9); Cell(1).PaddingVertical(2).AlignRight().Text(dl).FontSize(9); Cell(1).PaddingVertical(2).AlignRight().Text(da != 0 ? F(da) : "—").FontSize(9); }
                            Cell(6, "#D9E2F3").AlignRight().Text("الشركة: أورا ديفلوبرز إيجيبت للاستثمار العقاري").FontSize(11).Bold();
                            Info("الرقم الوظيفي", I("EMP_HR_CODE")); Info("الرقم الوظيفي ORA", I("ORA Code"));
                            Info("أسم الموظف", V("EMP_NAME_ALL")); Info("القسم", V("Payslip distribution")); Info("الشركة", V("EMP_COMPANY"));
                            RH();
                            DR("الراتب الاساسي", D("Basic Salary"), "التأمينات الاجتماعية – حصة الموظف", D("Social Insurance Employee Share"));
                            DR("الانتقالات", D("Transportation Allowance"), "الضريبة الشهرية", D("Monthly Tax"));
                            DR("بدل غلاء معيشة", D("Inflation Allowance"), "موبايل حد اقصي", 0);
                            DR("بدلات اخري", D("Housing Allowance"), "قيمة الايام المخصومة", D("UPL"));
                            DR("عدد ساعات اضافية", D("OT *1.5"), "بريميم", D("Premium"));
                            DR("قيمة الساعات الاضافية", D("Overtime Amount"), "ناوي باي", D("NowPay"));
                            DR("عدد الايام اضافية", D("National working Days"), "خصومات اخري", 0);
                            DR("قيمة الايام الاضافية", D("Holiday Over time Days"), "خصم صندوق إعانة الشهداء", D("Martyrs Fund"));
                            DR("بونص", 0, "الخصم الشهري للسلف من الراتب", D("Personal Loan Installment"));
                            DR("زيادة باثر رجعي", 0, "خصم مده تأمينة", 0);
                            DR("اخري", D("Previous Months"), "", 0);
                            var net = D("Net Salary");
                            Cell(2).PaddingVertical(5).AlignRight().Text("صافي الراتب").FontSize(14).Bold();
                            Cell(2).PaddingVertical(5).AlignRight().Text(F(net)).FontSize(14).Bold();
                            Cell(2).PaddingVertical(5);
                        });
                        col.Item().PaddingTop(10).AlignCenter().Text($"تم إنشاء هذا المستند بتاريخ {DateTime.Now:dd/MM/yyyy}").FontSize(7).FontColor("#999999");
                    });
                });
            }
        }).GeneratePdf(pdfPath);
    }

    private static string SafeFileName(string n) { var inv = Path.GetInvalidFileNameChars(); return new string(n.Where(c => !inv.Contains(c)).ToArray()); }
}
