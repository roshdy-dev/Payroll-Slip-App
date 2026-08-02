using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using PayrollSlipApp.Models;

namespace PayrollSlipApp.Services;

/// <summary>
/// Reads payroll data from an Excel (.xlsx) file.
/// The first row is treated as the header row containing column names.
/// If the first row has a yellow background, it's treated as the template marker row
/// (still used as headers -- the yellow formatting is informational only).
/// </summary>
public class ExcelReaderService
{
    /// <summary>
    /// Maps common Excel column header names (case-insensitive) to 
    /// the corresponding EmployeePayroll property for flexible column ordering.
    /// Add or adjust these mappings to match your organization's Excel header naming conventions.
    /// </summary>
    private static readonly Dictionary<string, string> ColumnMappings = new(StringComparer.OrdinalIgnoreCase)
    {
        // ── Identity columns ──
        ["employeeid"]          = nameof(EmployeePayroll.EmployeeId),
        ["employee id"]         = nameof(EmployeePayroll.EmployeeId),
        ["emp id"]              = nameof(EmployeePayroll.EmployeeId),
        ["empid"]               = nameof(EmployeePayroll.EmployeeId),
        ["payroll number"]      = nameof(EmployeePayroll.EmployeeId),
        ["staff id"]            = nameof(EmployeePayroll.EmployeeId),
        ["code"]                = nameof(EmployeePayroll.EmployeeId),
        ["emp_hr_code"]         = nameof(EmployeePayroll.EmployeeId),
        ["hr code"]             = nameof(EmployeePayroll.EmployeeId),
        ["hrcode"]              = nameof(EmployeePayroll.EmployeeId),
        ["ora code"]            = nameof(EmployeePayroll.EmployeeId),

        // ── Name columns ──
        ["employeename"]        = nameof(EmployeePayroll.EmployeeName),
        ["employee name"]       = nameof(EmployeePayroll.EmployeeName),
        ["name"]                = nameof(EmployeePayroll.EmployeeName),
        ["full name"]           = nameof(EmployeePayroll.EmployeeName),
        ["staff name"]          = nameof(EmployeePayroll.EmployeeName),
        ["emp_name_all"]        = nameof(EmployeePayroll.EmployeeName),

        // ── Department / Payslip Distribution columns ──
        // NOTE: Only "Payslip distribution" (the Excel column used for grouping payslips)
        // maps to EmployeePayroll.Department. Generic "Department" headers are intentionally
        // excluded to avoid conflicts when both columns exist in the same sheet.
        ["payslip distribution"]    = nameof(EmployeePayroll.Department),
        ["payslipdistribution"]     = nameof(EmployeePayroll.Department),

        // ── Job title columns ──
        ["jobtitle"]            = nameof(EmployeePayroll.JobTitle),
        ["job title"]           = nameof(EmployeePayroll.JobTitle),
        ["designation"]         = nameof(EmployeePayroll.JobTitle),
        ["position"]            = nameof(EmployeePayroll.JobTitle),
        ["title"]               = nameof(EmployeePayroll.JobTitle),
        ["emp_division"]        = nameof(EmployeePayroll.JobTitle),
        ["emp division"]        = nameof(EmployeePayroll.JobTitle),

        // ── Salary columns ──
        ["basicsalary"]         = nameof(EmployeePayroll.BasicSalary),
        ["basic salary"]        = nameof(EmployeePayroll.BasicSalary),
        ["basic"]               = nameof(EmployeePayroll.BasicSalary),
        ["base salary"]         = nameof(EmployeePayroll.BasicSalary),

        // ── Allowance columns ──
        ["housingallowance"]            = nameof(EmployeePayroll.HousingAllowance),
        ["housing allowance"]           = nameof(EmployeePayroll.HousingAllowance),
        ["housing"]                     = nameof(EmployeePayroll.HousingAllowance),
        ["transportationallowance"]     = nameof(EmployeePayroll.TransportationAllowance),
        ["transportation allowance"]    = nameof(EmployeePayroll.TransportationAllowance),
        ["transportation"]              = nameof(EmployeePayroll.TransportationAllowance),
        ["transport"]                   = nameof(EmployeePayroll.TransportationAllowance),
        ["otherallowances"]             = nameof(EmployeePayroll.OtherAllowances),
        ["other allowances"]            = nameof(EmployeePayroll.OtherAllowances),
        ["other allowance"]             = nameof(EmployeePayroll.OtherAllowances),
        ["bonus"]                       = nameof(EmployeePayroll.OtherAllowances),
        ["inflation allowance"]         = nameof(EmployeePayroll.InflationAllowance),
        ["inflationallowance"]          = nameof(EmployeePayroll.InflationAllowance),
        ["overtime amount"]             = nameof(EmployeePayroll.OtherAllowances),
        ["overtimeamount"]              = nameof(EmployeePayroll.OtherAllowances),
        ["other earning net amount"]    = nameof(EmployeePayroll.OtherAllowances),
        ["otherearningnetamount"]       = nameof(EmployeePayroll.OtherAllowances),

        // ── Deduction columns ──
        ["taxdeduction"]                = nameof(EmployeePayroll.TaxDeduction),
        ["tax deduction"]               = nameof(EmployeePayroll.TaxDeduction),
        ["tax"]                         = nameof(EmployeePayroll.TaxDeduction),
        ["income tax"]                  = nameof(EmployeePayroll.TaxDeduction),
        ["monthly tax"]                 = nameof(EmployeePayroll.TaxDeduction),
        ["monthlytax"]                  = nameof(EmployeePayroll.TaxDeduction),
        ["socialinsurance"]             = nameof(EmployeePayroll.SocialInsurance),
        ["social insurance"]            = nameof(EmployeePayroll.SocialInsurance),
        ["insurance"]                   = nameof(EmployeePayroll.SocialInsurance),
        ["national insurance"]          = nameof(EmployeePayroll.SocialInsurance),
        ["social insurance employee share"] = nameof(EmployeePayroll.SocialInsurance),
        ["socialinsuranceemployeeshare"]    = nameof(EmployeePayroll.SocialInsurance),
        ["otherdeductions"]             = nameof(EmployeePayroll.OtherDeductions),
        ["other deductions"]            = nameof(EmployeePayroll.OtherDeductions),
        ["loans"]                       = nameof(EmployeePayroll.OtherDeductions),
        ["advances"]                    = nameof(EmployeePayroll.OtherDeductions),
        ["hiring deduction"]            = nameof(EmployeePayroll.OtherDeductions),
        ["hiringdeduction"]             = nameof(EmployeePayroll.OtherDeductions),
        ["premium"]                     = nameof(EmployeePayroll.OtherDeductions),
        ["personal loan installment"]   = nameof(EmployeePayroll.OtherDeductions),
        ["personalloaninstallment"]     = nameof(EmployeePayroll.OtherDeductions),
        ["martyrs fund"]                = nameof(EmployeePayroll.OtherDeductions),
        ["martyrsfund"]                 = nameof(EmployeePayroll.OtherDeductions),
        ["nowpay"]                      = nameof(EmployeePayroll.OtherDeductions),
        ["now pay"]                     = nameof(EmployeePayroll.OtherDeductions),

        // ── Additional info columns ──
        ["payperiod"]               = nameof(EmployeePayroll.PayPeriod),
        ["pay period"]              = nameof(EmployeePayroll.PayPeriod),
        ["month"]                   = nameof(EmployeePayroll.PayPeriod),
        ["period"]                  = nameof(EmployeePayroll.PayPeriod),
        ["bankaccount"]             = nameof(EmployeePayroll.BankAccount),
        ["bank account"]            = nameof(EmployeePayroll.BankAccount),
        ["account number"]          = nameof(EmployeePayroll.BankAccount),
        ["accountnumber"]           = nameof(EmployeePayroll.BankAccount),
        ["accountno"]               = nameof(EmployeePayroll.BankAccount),
        ["bank"]                    = nameof(EmployeePayroll.BankAccount),
        ["joiningdate"]             = nameof(EmployeePayroll.JoiningDate),
        ["joining date"]            = nameof(EmployeePayroll.JoiningDate),
        ["date of joining"]         = nameof(EmployeePayroll.JoiningDate),
        ["hire date"]               = nameof(EmployeePayroll.JoiningDate),
        ["emp_company"]             = nameof(EmployeePayroll.JoiningDate),
        ["emp company"]             = nameof(EmployeePayroll.JoiningDate),
    };

    /// <summary>
    /// Reads all employee payroll records from the specified Excel file.
    /// </summary>
    /// <param name="filePath">Full path to the .xlsx file.</param>
    /// <returns>A list of EmployeePayroll objects parsed from the spreadsheet.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown when required columns (EmployeeId, Department, EmployeeName) are missing.
    /// </exception>
    public List<EmployeePayroll> ReadPayrollData(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Excel file not found: {filePath}");

        var employees = new List<EmployeePayroll>();

        // Open the workbook in read-only mode for performance
        using var workbook = new XLWorkbook(filePath);

        // Find the worksheet that contains data headers (not the template/layout sheet)
        var worksheet = FindDataWorksheet(workbook);

        // --- Step 1: Read the header row (row 1) ---
        // Build a mapping: ColumnIndex -> PropertyName
        var headerRow = worksheet.Row(1);
        var columnMap = new Dictionary<int, string>(); // column number -> EmployeePayroll property name
        var rawHeaders = new Dictionary<int, string>(); // column number -> normalized header (for RawData)

        for (int col = 1; col <= headerRow.LastCellUsed()?.Address.ColumnNumber; col++)
        {
            var headerText = headerRow.Cell(col).GetString().Trim();

            // Skip empty header cells
            if (string.IsNullOrWhiteSpace(headerText))
                continue;

            // Store normalized header for raw data lookup
            var normHeader = NormalizeKey(headerText);
            rawHeaders[col] = normHeader;

            // Try to match the header text to a known property
            if (ColumnMappings.TryGetValue(headerText, out var propertyName))
            {
                columnMap[col] = propertyName;
            }
            else
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ExcelReader] Unknown column header '{headerText}' at column {col} -- will be ignored.");
            }
        }

        // --- Step 2: Validate that we have the minimum required columns ---
        var mappedProperties = columnMap.Values.ToHashSet();
        var foundHeaders = columnMap.Keys.Select(c => headerRow.Cell(c).GetString().Trim()).ToList();
        var allHeaders = $"Found headers: [{string.Join(", ", foundHeaders)}]";

        if (!mappedProperties.Contains(nameof(EmployeePayroll.EmployeeId)))
            throw new InvalidOperationException(
                $"Required column 'Employee ID' not found in the Excel header row. {allHeaders}. " +
                "Please ensure a column header matches one of: EMP_HR_CODE, EmployeeID, Emp ID, Staff ID, Code.");

        if (!mappedProperties.Contains(nameof(EmployeePayroll.EmployeeName)))
            throw new InvalidOperationException(
                $"Required column 'Employee Name' not found in the Excel header row. {allHeaders}. " +
                "Please ensure a column header matches one of: EMP_NAME_ALL, Employee Name, Name, Full Name.");

        // --- Step 3: Read data rows (starting from row 2) ---
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

        for (int row = 2; row <= lastRow; row++)
        {
            var dataRow = worksheet.Row(row);

            // Skip completely empty rows
            if (dataRow.IsEmpty())
                continue;

            var employee = new EmployeePayroll();

            // Populate RawData: normalized header → cell value (for template matching)
            for (int col = 1; col <= headerRow.LastCellUsed()?.Address.ColumnNumber; col++)
            {
                if (rawHeaders.TryGetValue(col, out var normKey))
                {
                    var rawVal = dataRow.Cell(col).GetString().Trim();
                    if (!string.IsNullOrWhiteSpace(rawVal))
                        employee.RawData[normKey] = rawVal;
                }
            }

            // Populate properties based on the column mapping
            foreach (var (colIndex, propName) in columnMap)
            {
                var cellValue = dataRow.Cell(colIndex).GetString().Trim();

                // Skip empty cells
                if (string.IsNullOrWhiteSpace(cellValue))
                    continue;

                SetPropertyValue(employee, propName, cellValue);
            }

            // Only add if EmployeeId is populated (skip blank rows)
            if (!string.IsNullOrWhiteSpace(employee.EmployeeId))
            {
                employees.Add(employee);
            }
        }

        // ── Set default PayPeriod for any employee missing it ──
        // Uses Arabic month names matching the template convention
        var defaultPeriod = GetDefaultPayPeriod();
        foreach (var emp in employees)
        {
            if (string.IsNullOrWhiteSpace(emp.PayPeriod))
                emp.PayPeriod = defaultPeriod;
        }

        return employees;
    }

    /// <summary>
    /// Finds the worksheet that contains actual payroll data headers.
    /// Scans sheets in order, looking for the first sheet whose row 1 contains
    /// at least one recognized Employee ID column header.
    /// Falls back to the first worksheet if no data sheet is detected.
    /// </summary>
    private static IXLWorksheet FindDataWorksheet(XLWorkbook workbook)
    {
        // Known Employee ID header patterns (same as ColumnMappings for employeeid)
        var idHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "employeeid", "employee id", "emp id", "empid", "payroll number",
            "staff id", "code", "emp_hr_code", "hr code", "hrcode", "ora code"
        };

        foreach (var ws in workbook.Worksheets)
        {
            var headerRow = ws.Row(1);
            var lastCol = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;

            for (int col = 1; col <= lastCol; col++)
            {
                var headerText = headerRow.Cell(col).GetString().Trim();
                if (!string.IsNullOrWhiteSpace(headerText) && idHeaders.Contains(headerText))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[ExcelReader] Using sheet '{ws.Name}' (found ID header '{headerText}' at col {col}).");
                    return ws;
                }
            }
        }

        // Fallback: use the first worksheet
        System.Diagnostics.Debug.WriteLine(
            "[ExcelReader] No data sheet detected by headers, falling back to first worksheet.");
        return workbook.Worksheets.First();
    }

    /// <summary>
    /// Returns a default pay period string in Arabic (e.g., "يوليو 2026").
    /// </summary>
    private static string GetDefaultPayPeriod()
    {
        var now = DateTime.Now;
        string[] arabicMonths = { "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
                                   "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر" };
        return $"{arabicMonths[now.Month - 1]} {now.Year}";
    }

    /// <summary>
    /// Normalizes a string key: lowercase, letters+digits only.
    /// Same function used in WordGenerator/PdfGenerator for matching.
    /// </summary>
    internal static string NormalizeKey(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return new string(s.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
    }

    /// <summary>
    /// Groups employees by their Department property.
    /// Uses the same logic as GroupByColumn with "Payslip distribution".
    /// </summary>
    /// <param name="employees">Flat list of employees.</param>
    /// <returns>List of DepartmentGroup, one per unique department.</returns>
    public List<DepartmentGroup> GroupByDepartment(List<EmployeePayroll> employees)
    {
        return GroupByColumn(employees, "Payslip distribution");
    }

    /// <summary>
    /// Groups employees by the value of a specific Excel column (from RawData).
    /// Each unique value in that column becomes its own output group (Word/PDF file).
    /// 
    /// The column name is normalized (lowercase, letters+digits only) before lookup
    /// in RawData, matching how ExcelReaderService stores headers.
    /// </summary>
    /// <param name="employees">Flat list of employees with RawData populated.</param>
    /// <param name="columnHeader">
    /// Exact Excel column header name to group by (e.g., "Payslip distribution", "Company", "Department").
    /// Case-insensitive, spaces/punctuation are stripped for matching.
    /// </param>
    /// <returns>List of DepartmentGroup, one per unique column value.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown when the specified column is not found in any employee's RawData.
    /// </exception>
    public List<DepartmentGroup> GroupByColumn(List<EmployeePayroll> employees, string columnHeader)
    {
        if (string.IsNullOrWhiteSpace(columnHeader))
            throw new ArgumentException("Column header name cannot be empty.", nameof(columnHeader));

        if (employees == null || employees.Count == 0)
            return new List<DepartmentGroup>();

        var normalizedColumn = NormalizeKey(columnHeader);

        // Validate that the column exists in at least one employee's RawData
        var hasColumn = employees.Any(e => e.RawData.ContainsKey(normalizedColumn));
        if (!hasColumn)
        {
            var availableColumns = employees
                .SelectMany(e => e.RawData.Keys)
                .Distinct()
                .OrderBy(k => k)
                .ToList();

            throw new InvalidOperationException(
                $"The column separator '{columnHeader}' (normalized: '{normalizedColumn}') " +
                $"was not found in the Excel data. Available columns: [{string.Join(", ", availableColumns)}]. " +
                "Please check the 'ColumnSeparator' value in AppConfig.json and ensure it matches " +
                "one of the Excel column headers exactly.");
        }

        // Group employees by the raw value of the specified column
        return employees
            .GroupBy(e =>
            {
                // Try RawData first (the normalized key), then fall back to Department property
                if (e.RawData.TryGetValue(normalizedColumn, out var value))
                    return value;

                // Fallback: try matching known properties
                if (normalizedColumn == NormalizeKey("Payslip distribution"))
                    return e.Department;

                return "(Unknown)";
            })
            .Select(g => new DepartmentGroup
            {
                DepartmentName = g.Key,
                Employees = g.ToList()
            })
            .ToList();
    }

    /// <summary>
    /// Sets a property value on an EmployeePayroll instance by property name.
    /// Uses reflection for flexibility with the dynamic column mapping.
    /// </summary>
    private static void SetPropertyValue(EmployeePayroll employee, string propertyName, string rawValue)
    {
        var property = typeof(EmployeePayroll).GetProperty(propertyName);
        if (property == null) return;

        // Determine the target type and convert the string value accordingly
        if (property.PropertyType == typeof(decimal))
        {
            // Parse decimal values -- handle various formats ($1,234.56, 1234,56, etc.)
            if (decimal.TryParse(rawValue.Replace("$", "").Replace(",", "").Replace(" ", ""),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var decimalValue))
            {
                property.SetValue(employee, decimalValue);
            }
        }
        else if (property.PropertyType == typeof(string))
        {
            property.SetValue(employee, rawValue);
        }
    }
}
