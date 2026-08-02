namespace PayrollSlipApp.Models;

/// <summary>
/// Application configuration loaded from AppConfig.json.
/// </summary>
public class AppConfig
{
    /// <summary>
    /// The Excel column header name used to separate/split the output files.
    /// Each unique value in this column gets its own Word/PDF payslip file.
    /// Default: "Payslip distribution"
    /// </summary>
    public string ColumnSeparator { get; set; } = "Payslip distribution";

    /// <summary>
    /// Column-level data type overrides for value formatting when filling the template.
    /// If a column is not listed here, the default behaviour is:
    ///   - If the value parses as decimal → formatted with thousand separators + 2 decimal places
    ///   - Otherwise → displayed as-is (string)
    /// </summary>
    public List<FormatDataType> FormatDataTypes { get; set; } = new();
}

/// <summary>
/// Defines how a specific column/placeholder value should be formatted in the output.
/// </summary>
public class FormatDataType
{
    /// <summary>
    /// The column header name in Excel, or placeholder name in the Word template
    /// (e.g., "Month", "Year" for pay-period placeholders).
    /// Case-insensitive match after normalization (spaces/punctuation stripped).
    /// </summary>
    public string ColumnName { get; set; } = string.Empty;

    /// <summary>
    /// One of: integer, decimal, string, date, dateTime, PayPeriod_Month, PayPeriod_Year.
    /// - integer          : plain number, no thousand separator, no decimals (e.g., "100001")
    /// - decimal          : thousand separators + 2 decimal places (e.g., "7,409.00")
    /// - string           : displayed as-is, no numeric formatting
    /// - date             : formatted as dd/MM/yyyy
    /// - dateTime         : formatted as dd/MM/yyyy HH:mm
    /// - PayPeriod_Month  : replaced with the selected pay-period month name
    /// - PayPeriod_Year   : replaced with the selected pay-period year
    /// </summary>
    public string DataType { get; set; } = "string";

    /// <summary>
    /// Language for PayPeriod_Month formatting. "Arabic" (default) or "English".
    /// Ignored for all other data types.
    /// </summary>
    public string Language { get; set; } = "Arabic";
}
