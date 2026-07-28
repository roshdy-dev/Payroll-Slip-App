using System.Collections.Generic;

namespace PayrollSlipApp.Models;

/// <summary>
/// Represents a single employee's payroll data extracted from the Excel file.
/// Each row in the Excel (after the header) maps to one EmployeePayroll instance.
/// </summary>
public class EmployeePayroll
{
    /// <summary>Employee ID or payroll number (e.g., "EMP001").</summary>
    public string EmployeeId { get; set; } = string.Empty;

    /// <summary>Full name of the employee.</summary>
    public string EmployeeName { get; set; } = string.Empty;

    /// <summary>Department name (e.g., "Engineering", "Sales"). Used to group payslips.</summary>
    public string Department { get; set; } = string.Empty;

    /// <summary>Job title or designation.</summary>
    public string JobTitle { get; set; } = string.Empty;

    /// <summary>Basic monthly salary before additions/deductions.</summary>
    public decimal BasicSalary { get; set; }

    /// <summary>Housing allowance amount.</summary>
    public decimal HousingAllowance { get; set; }

    /// <summary>Transportation allowance amount.</summary>
    public decimal TransportationAllowance { get; set; }

    /// <summary>Inflation / cost-of-living allowance.</summary>
    public decimal InflationAllowance { get; set; }

    /// <summary>Any other allowances or bonuses not covered above.</summary>
    public decimal OtherAllowances { get; set; }

    /// <summary>Total gross earnings (Basic + Housing + Transport + Inflation + Other).</summary>
    public decimal TotalEarnings => BasicSalary + HousingAllowance + TransportationAllowance + InflationAllowance + OtherAllowances;

    /// <summary>Income tax deducted.</summary>
    public decimal TaxDeduction { get; set; }

    /// <summary>Social insurance / National insurance contribution.</summary>
    public decimal SocialInsurance { get; set; }

    /// <summary>Any other deductions (loans, advances, etc.).</summary>
    public decimal OtherDeductions { get; set; }

    /// <summary>Total deductions (Tax + Social Insurance + Others).</summary>
    public decimal TotalDeductions => TaxDeduction + SocialInsurance + OtherDeductions;

    /// <summary>Net salary payable = TotalEarnings - TotalDeductions.</summary>
    public decimal NetSalary => TotalEarnings - TotalDeductions;

    /// <summary>Month/year this payslip covers (e.g., "July 2026").</summary>
    public string PayPeriod { get; set; } = string.Empty;

    /// <summary>Bank account number for salary transfer.</summary>
    public string BankAccount { get; set; } = string.Empty;

    /// <summary>Date of joining the company.</summary>
    public string JoiningDate { get; set; } = string.Empty;

    /// <summary>
    /// Raw Excel data: normalized column header → cell value.
    /// Filled by ExcelReaderService. Used by Word/Pdf generators for direct template matching.
    /// </summary>
    public Dictionary<string, string> RawData { get; set; } = new();
}

/// <summary>
/// Groups employees by department. One Word/PDF file is generated 
/// per department containing all its employees' payslips.
/// </summary>
public class DepartmentGroup
{
    /// <summary>Department name.</summary>
    public string DepartmentName { get; set; } = string.Empty;

    /// <summary>All employees belonging to this department.</summary>
    public List<EmployeePayroll> Employees { get; set; } = new();
}
