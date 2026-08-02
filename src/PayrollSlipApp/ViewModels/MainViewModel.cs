using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PayrollSlipApp.Models;
using PayrollSlipApp.Services;

namespace PayrollSlipApp.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    private readonly ExcelReaderService _excelReader;
    private readonly WordGeneratorService _wordGenerator;
    private readonly PdfGeneratorService _pdfGenerator;
    private readonly AppConfig _appConfig;

    [ObservableProperty]
    private string _selectedFilePath = string.Empty;

    [ObservableProperty]
    private string _selectedFileName = "No file selected";

    /// <summary>Whether a file has been selected.</summary>
    public bool HasFile => !string.IsNullOrWhiteSpace(SelectedFilePath);

    /// <summary>Whether generation can start (file selected + not already processing).</summary>
    public bool CanGenerate => HasFile && !IsProcessing;

    [ObservableProperty]
    private string _statusMessage = "Ready — select an Excel payroll file to begin";

    [ObservableProperty]
    private bool _isProcessing;

    [ObservableProperty]
    private int _progress;

    [ObservableProperty]
    private bool _isProgressVisible;

    [ObservableProperty]
    private string _outputDirectory = string.Empty;

    [ObservableProperty]
    private bool _canOpenDirectory;

    [ObservableProperty]
    private string _logMessages = string.Empty;

    private string _lastFormat = string.Empty;

    // ── Month / Year Picker ──

    private static readonly string[] ArabicMonths =
        { "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
          "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر" };

    /// <summary>Available month names (filtered by year selection).</summary>
    public List<string> MonthNames { get; } = new();

    /// <summary>Available years (past 5 years + current).</summary>
    public List<int> AvailableYears { get; }

    [ObservableProperty]
    private int _selectedYear;

    [ObservableProperty]
    private int _selectedMonthIndex = -1; // -1 = no selection

    /// <summary>The pay period string in Arabic: "يوليو 2026"</summary>
    public string PayPeriod =>
        SelectedMonthIndex >= 0 && SelectedMonthIndex < 12
            ? $"{ArabicMonths[SelectedMonthIndex]} {SelectedYear}"
            : "";

    partial void OnSelectedYearChanged(int value)
    {
        RefreshMonthList();
    }

    private void RefreshMonthList()
    {
        MonthNames.Clear();
        var now = DateTime.Now;
        int maxMonth = (SelectedYear == now.Year) ? now.Month : 12;
        for (int m = 1; m <= maxMonth; m++)
            MonthNames.Add(ArabicMonths[m - 1]);

        // Reset month selection if it's now out of range
        if (SelectedMonthIndex >= MonthNames.Count)
            SelectedMonthIndex = MonthNames.Count > 0 ? MonthNames.Count - 1 : -1;
        else if (SelectedMonthIndex < 0 && MonthNames.Count > 0)
            SelectedMonthIndex = MonthNames.Count - 1; // default to latest

        OnPropertyChanged(nameof(PayPeriod));
    }

    public MainViewModel()
    {
        _excelReader = new ExcelReaderService();
        _appConfig = LoadAppConfig();
        _wordGenerator = new WordGeneratorService(_appConfig);
        _pdfGenerator = new PdfGeneratorService(_appConfig);

        // Populate years: current year and 5 past years
        var now = DateTime.Now;
        AvailableYears = Enumerable.Range(now.Year - 5, 6).Reverse().ToList();
        SelectedYear = now.Year;
        RefreshMonthList();
    }

    /// <summary>
    /// Loads AppConfig.json from next to the executable, or from the project root during development.
    /// Falls back to default settings if the file is missing or malformed.
    /// </summary>
    private static AppConfig LoadAppConfig()
    {
        // Resolve config path dynamically (same strategy as WordGeneratorService.TemplatePath)
        var exeDir = AppDomain.CurrentDomain.BaseDirectory;
        var localConfig = Path.Combine(exeDir, "AppConfig.json");
        var projectConfig = @"D:\Payroll Slip App\AppConfig.json";

        string? configPath = null;
        if (File.Exists(localConfig))
            configPath = localConfig;
        else if (File.Exists(projectConfig))
            configPath = projectConfig;

        if (configPath == null)
        {
            System.Diagnostics.Debug.WriteLine("[MainViewModel] AppConfig.json not found, using defaults.");
            return new AppConfig();
        }

        try
        {
            var json = File.ReadAllText(configPath);
            var config = JsonSerializer.Deserialize<AppConfig>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
            if (config != null && !string.IsNullOrWhiteSpace(config.ColumnSeparator))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[MainViewModel] Loaded AppConfig: ColumnSeparator='{config.ColumnSeparator}'");
                return config;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[MainViewModel] Failed to parse AppConfig.json: {ex.Message}");
        }

        return new AppConfig();
    }

    /// <summary>
    /// Opens a file picker dialog filtered to Excel files (.xlsx, .xls).
    /// Tries the modern StorageProvider API first, falls back to legacy dialog.
    /// </summary>
    [RelayCommand]
    private async Task BrowseFileAsync()
    {
        try
        {
            // Get the main window reference (cross-platform safe)
            var mainWindow = (Avalonia.Application.Current?.ApplicationLifetime as
                Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime)?.MainWindow;

            if (mainWindow == null)
            {
                AppendLog("❌ Could not access application window.");
                return;
            }

            string? selectedPath = null;

            // --- Strategy 1: Try modern StorageProvider API ---
            try
            {
                var storageProvider = mainWindow.StorageProvider;
                if (storageProvider != null && storageProvider.CanOpen)
                {
                    var files = await storageProvider.OpenFilePickerAsync(
                        new Avalonia.Platform.Storage.FilePickerOpenOptions
                        {
                            Title = "Select Payroll Excel File",
                            AllowMultiple = false,
                            FileTypeFilter = new List<Avalonia.Platform.Storage.FilePickerFileType>
                            {
                                new("Excel Files") { Patterns = new[] { "*.xlsx", "*.xls" } },
                                new("All Files")  { Patterns = new[] { "*.*" } }
                            }
                        });

                    if (files != null && files.Count > 0)
                        selectedPath = files[0].Path.LocalPath;
                }
            }
            catch
            {
                // StorageProvider failed — fall back to legacy dialog below
            }

            // --- Strategy 2: Fall back to legacy OpenFileDialog ---
            if (string.IsNullOrEmpty(selectedPath))
            {
                var dialog = new Avalonia.Controls.OpenFileDialog
                {
                    Title = "Select Payroll Excel File",
                    AllowMultiple = false,
                    Filters = new List<Avalonia.Controls.FileDialogFilter>
                    {
                        new() { Name = "Excel Files", Extensions = new List<string> { "xlsx", "xls" } },
                        new() { Name = "All Files",   Extensions = new List<string> { "*" } }
                    }
                };

                var result = await dialog.ShowAsync(mainWindow);
                if (result != null && result.Length > 0)
                    selectedPath = result[0];
            }

            // --- Process the selected file ---
            SetSelectedFile(selectedPath);
        }
        catch (Exception ex)
        {
            AppendLog($"❌ Error opening file dialog: {ex.Message}");
            AppendLog($"   Details: {ex.InnerException?.Message}");
        }
    }

    /// <summary>Handles drag-and-drop file selection from the UI.</summary>
    public void OnFileDropped(string filePath)
    {
        if (!IsProcessing && File.Exists(filePath))
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".xlsx" || ext == ".xls")
                SetSelectedFile(filePath);
            else
                AppendLog("⚠️ Please drop an Excel file (.xlsx or .xls).");
        }
    }

    /// <summary>Sets the selected file and updates related UI state.</summary>
    private void SetSelectedFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
        SelectedFilePath = path;
        SelectedFileName = Path.GetFileName(path);
        OnPropertyChanged(nameof(HasFile));
        OnPropertyChanged(nameof(CanGenerate));
        AppendLog($"📁 {SelectedFileName}");
    }

    /// <summary>Clears the currently selected file.</summary>
    [RelayCommand]
    private void ClearFile()
    {
        SelectedFilePath = string.Empty;
        SelectedFileName = "No file selected";
        OnPropertyChanged(nameof(HasFile));
        OnPropertyChanged(nameof(CanGenerate));
    }

    /// <summary>Notifies HasFile/CanGenerate when SelectedFilePath changes.</summary>
    partial void OnSelectedFilePathChanged(string value)
    {
        OnPropertyChanged(nameof(HasFile));
        OnPropertyChanged(nameof(CanGenerate));
    }

    partial void OnIsProcessingChanged(bool value)
    {
        OnPropertyChanged(nameof(CanGenerate));
    }

    /// <summary>
    /// Generates Word (.docx) payslips from the selected Excel file.
    /// </summary>
    [RelayCommand]
    private async Task GenerateWordAsync()
    {
        _lastFormat = "Word";
        await GeneratePayslipsAsync("Word");
    }

    /// <summary>
    /// Generates PDF payslips from the selected Excel file.
    /// </summary>
    [RelayCommand]
    private async Task GeneratePdfAsync()
    {
        _lastFormat = "PDF";
        await GeneratePayslipsAsync("PDF");
    }

    /// <summary>
    /// Generates BOTH Word and PDF payslips from the selected Excel file.
    /// </summary>
    [RelayCommand]
    private async Task GenerateBothAsync()
    {
        _lastFormat = "Both";
        await GeneratePayslipsAsync("Both");
    }

    /// <summary>
    /// Core generation logic shared by all format options.
    /// </summary>
    private async Task GeneratePayslipsAsync(string format)
    {
        // ── Validation ──
        if (string.IsNullOrWhiteSpace(SelectedFilePath))
        {
            StatusMessage = "⚠️ Please select an Excel file first.";
            return;
        }

        if (!File.Exists(SelectedFilePath))
        {
            StatusMessage = "⚠️ The selected file no longer exists. Please re-select.";
            SelectedFilePath = string.Empty;
            SelectedFileName = "No file selected";
            return;
        }

        // ── Reset UI state ──
        IsProcessing = true;
        IsProgressVisible = true;
        Progress = 0;
        LogMessages = string.Empty;
        CanOpenDirectory = false;

        AppendLog("═══════════════════════════════════════");
        AppendLog($"Starting {format} payslip generation...");
        AppendLog("═══════════════════════════════════════");

        try
        {
            // Run the heavy work on a background thread to keep UI responsive
            await Task.Run(() =>
            {
                // ── Step 1: Read Excel (20%) ──
                UpdateProgress(10);
                AppendLog("📖 Reading Excel file...");

                var employees = _excelReader.ReadPayrollData(SelectedFilePath);

                // ── Set pay period from UI picker ──
                var payPeriod = PayPeriod;
                foreach (var e in employees)
                    e.PayPeriod = payPeriod;

                UpdateProgress(20);
                AppendLog($"✅ Read {employees.Count} employee records.");

                if (employees.Count == 0)
                {
                    AppendLog("⚠️ No employee records found in the Excel file.");
                    return;
                }

                // ── Step 2: Group by the configured column separator (25%) ──
                UpdateProgress(25);
                var separator = _appConfig.ColumnSeparator;
                AppendLog($"📊 Grouping by column: '{separator}'");
                var departments = _excelReader.GroupByColumn(employees, separator);
                AppendLog($"📊 Found {departments.Count} department(s):");
                foreach (var dept in departments)
                {
                    AppendLog($"   • {dept.DepartmentName}: {dept.Employees.Count} employee(s)");
                }

                // ── Step 3: Create output directory (30%) ──
                UpdateProgress(30);

                // Output goes to a subfolder next to the Excel file
                var excelDir = Path.GetDirectoryName(SelectedFilePath)!;
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                OutputDirectory = Path.Combine(excelDir, $"Payslips_{timestamp}");
                Directory.CreateDirectory(OutputDirectory);

                AppendLog($"📂 Output directory: {OutputDirectory}");

                // ── Step 4: Generate documents ──
                var generatedFiles = new List<string>();

                if (format == "Word" || format == "Both")
                {
                    UpdateProgress(40);
                    AppendLog("📝 Generating Word documents...");
                    var wordFiles = _wordGenerator.GenerateWordDocuments(departments, OutputDirectory);
                    generatedFiles.AddRange(wordFiles);
                    UpdateProgress(65);
                    AppendLog($"✅ Generated {wordFiles.Count} Word file(s).");
                    foreach (var f in wordFiles)
                        AppendLog($"   📄 {Path.GetFileName(f)}");
                }

                if (format == "PDF" || format == "Both")
                {
                    var startProgress = format == "Both" ? 65 : 40;
                    UpdateProgress(startProgress);
                    AppendLog("📄 Generating PDF documents...");
                    var pdfFiles = _pdfGenerator.GeneratePdfDocuments(departments, OutputDirectory);
                    generatedFiles.AddRange(pdfFiles);
                    UpdateProgress(format == "Both" ? 90 : 65);
                    AppendLog($"✅ Generated {pdfFiles.Count} PDF file(s).");
                    foreach (var f in pdfFiles)
                        AppendLog($"   📑 {Path.GetFileName(f)}");

                    // PDF-only mode: clean up intermediate Word files
                    if (format == "PDF")
                    {
                        foreach (var f in Directory.GetFiles(OutputDirectory, "*.docx"))
                        {
                            try { File.Delete(f); } catch { }
                        }
                    }
                }

                UpdateProgress(100);
                AppendLog("═══════════════════════════════════════");
                AppendLog($"🎉 All done! {generatedFiles.Count} file(s) generated successfully.");
                AppendLog($"📍 Location: {OutputDirectory}");
                AppendLog("═══════════════════════════════════════");
            });

            CanOpenDirectory = true;
            StatusMessage = $"✅ {format} payslips generated successfully!";
        }
        catch (Exception ex)
        {
            AppendLog($"❌ ERROR: {ex.Message}");
            if (ex.InnerException != null)
                AppendLog($"   Details: {ex.InnerException.Message}");
            StatusMessage = $"❌ Generation failed: {ex.Message}";
        }
        finally
        {
            IsProcessing = false;
        }
    }

    /// <summary>
    /// Opens the output directory in the system's file explorer.
    /// </summary>
    [RelayCommand]
    private void OpenOutputDirectory()
    {
        if (string.IsNullOrWhiteSpace(OutputDirectory) || !Directory.Exists(OutputDirectory))
        {
            StatusMessage = "⚠️ Output directory not available.";
            return;
        }

        try
        {
            // Cross-platform directory opening
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Process.Start("explorer.exe", OutputDirectory);
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                Process.Start("open", OutputDirectory);
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                Process.Start("xdg-open", OutputDirectory);
            }

            AppendLog($"📂 Opened: {OutputDirectory}");
        }
        catch (Exception ex)
        {
            AppendLog($"❌ Cannot open directory: {ex.Message}");
            StatusMessage = $"❌ Cannot open directory: {ex.Message}";
        }
    }

    // ─── Helpers ───

    private void UpdateProgress(int value)
    {
        // Marshal back to UI thread
        Avalonia.Threading.Dispatcher.UIThread.Post(() =>
        {
            Progress = value;
        });
    }

    private void AppendLog(string message)
    {
        Avalonia.Threading.Dispatcher.UIThread.Post(() =>
        {
            LogMessages += message + Environment.NewLine;
        });
    }
}
