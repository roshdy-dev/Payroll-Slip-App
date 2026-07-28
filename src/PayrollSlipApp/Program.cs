using System;
using Avalonia;

namespace PayrollSlipApp;

/// <summary>
/// Program entry point.
/// Builds and starts the Avalonia desktop application.
/// 
/// Avalonia handles cross-platform rendering automatically —
/// the same code runs on Windows, macOS, and Linux.
/// </summary>
internal class Program
{
    // STAThread is required for Avalonia on all desktop platforms
    [STAThread]
    public static void Main(string[] args)
    {
        BuildAvaloniaApp()
            .StartWithClassicDesktopLifetime(args);
    }

    /// <summary>
    /// Configures the Avalonia application builder.
    /// </summary>
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()      // Auto-detect OS (Win, Mac, Linux)
            .WithInterFont()           // Use Inter font family
            .LogToTrace();             // Enable debug logging
}
