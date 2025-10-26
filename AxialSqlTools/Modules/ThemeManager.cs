using System;
using System.Windows;
using System.Windows.Media;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.PlatformUI;
using NLog;

namespace AxialSqlTools
{
    /// <summary>
    /// Manages theme colors for SSMS dark/light mode compatibility
    /// </summary>
    public static class ThemeManager
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        
        private static readonly Color DarkBackground = Color.FromRgb(30, 30, 30);
        private static readonly Color DarkForeground = Color.FromRgb(220, 220, 220);
        private static readonly Color DarkHeaderBackground = Color.FromRgb(45, 45, 48);
        
        private static readonly Color LightBackground = Color.FromRgb(255, 255, 255);
        private static readonly Color LightForeground = Color.FromRgb(0, 0, 0);
        private static readonly Color LightHeaderBackground = Color.FromRgb(240, 240, 240);
        
        // Enable diagnostic mode to show message boxes with theme info
        // Set to false after initial testing to avoid popup spam
        public static bool DiagnosticMode { get; set; } = false;

        /// <summary>
        /// Determines if the current theme is dark
        /// </summary>
        public static bool IsDarkTheme()
        {
            try
            {
                // Try to get the VS theme colors
                var backgroundColor = VSColorTheme.GetThemedColor(EnvironmentColors.ToolWindowBackgroundColorKey);
                
                // Calculate luminance to determine if it's dark
                double luminance = (0.299 * backgroundColor.R + 0.587 * backgroundColor.G + 0.114 * backgroundColor.B) / 255;
                
                bool isDark = luminance < 0.5;
                
                string diagnosticMsg = $"Theme Detection:\n" +
                    $"Background Color: R={backgroundColor.R}, G={backgroundColor.G}, B={backgroundColor.B}\n" +
                    $"Luminance: {luminance:F3}\n" +
                    $"Is Dark Theme: {isDark}";
                
                Logger.Info(diagnosticMsg);
                
                if (DiagnosticMode)
                {
                    MessageBox.Show(diagnosticMsg, "ThemeManager - IsDarkTheme()", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                
                return isDark;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Failed to detect theme: {ex.Message}\nDefaulting to light theme";
                Logger.Error(ex, "Theme detection failed");
                
                if (DiagnosticMode)
                {
                    MessageBox.Show(errorMsg, "ThemeManager - IsDarkTheme() ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                
                // Default to light theme if we can't detect
                return false;
            }
        }

        /// <summary>
        /// Gets the appropriate background color for the current theme
        /// </summary>
        public static Color GetBackgroundColor()
        {
            try
            {
                var color = VSColorTheme.GetThemedColor(EnvironmentColors.ToolWindowBackgroundColorKey);
                var result = Color.FromRgb(color.R, color.G, color.B);
                
                Logger.Info($"GetBackgroundColor: R={result.R}, G={result.G}, B={result.B}");
                
                return result;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Failed to get background color, using fallback");
                return IsDarkTheme() ? DarkBackground : LightBackground;
            }
        }

        /// <summary>
        /// Gets the appropriate foreground (text) color for the current theme
        /// </summary>
        public static Color GetForegroundColor()
        {
            try
            {
                var color = VSColorTheme.GetThemedColor(EnvironmentColors.ToolWindowTextColorKey);
                var result = Color.FromRgb(color.R, color.G, color.B);
                
                Logger.Info($"GetForegroundColor: R={result.R}, G={result.G}, B={result.B}");
                
                return result;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Failed to get foreground color, using fallback");
                return IsDarkTheme() ? DarkForeground : LightForeground;
            }
        }

        /// <summary>
        /// Gets the appropriate header background color for the current theme
        /// </summary>
        public static Color GetHeaderBackgroundColor()
        {
            try
            {
                var color = VSColorTheme.GetThemedColor(EnvironmentColors.CommandBarGradientBeginColorKey);
                return Color.FromRgb(color.R, color.G, color.B);
            }
            catch
            {
                return IsDarkTheme() ? DarkHeaderBackground : LightHeaderBackground;
            }
        }

        /// <summary>
        /// Applies theme colors to a FrameworkElement
        /// </summary>
        public static void ApplyTheme(FrameworkElement element)
        {
            if (element == null) return;

            try
            {
                var bgColor = GetBackgroundColor();
                var fgColor = GetForegroundColor();
                var bgBrush = new SolidColorBrush(bgColor);
                var fgBrush = new SolidColorBrush(fgColor);

                string elementInfo = $"Applying theme to: {element.GetType().Name}";
                Logger.Info(elementInfo);

                // Try to set Background and Foreground properties
                var bgProperty = element.GetType().GetProperty("Background");
                if (bgProperty != null && bgProperty.CanWrite)
                {
                    bgProperty.SetValue(element, bgBrush);
                    Logger.Info($"  Background set: R={bgColor.R}, G={bgColor.G}, B={bgColor.B}");
                }
                else
                {
                    Logger.Warn($"  Background property not writable on {element.GetType().Name}");
                }

                var fgProperty = element.GetType().GetProperty("Foreground");
                if (fgProperty != null && fgProperty.CanWrite)
                {
                    fgProperty.SetValue(element, fgBrush);
                    Logger.Info($"  Foreground set: R={fgColor.R}, G={fgColor.G}, B={fgColor.B}");
                }
                else
                {
                    Logger.Warn($"  Foreground property not writable on {element.GetType().Name}");
                }
                
                if (DiagnosticMode)
                {
                    string diagnosticMsg = $"Applied Theme to {element.GetType().Name}:\n" +
                        $"Background: R={bgColor.R}, G={bgColor.G}, B={bgColor.B}\n" +
                        $"Foreground: R={fgColor.R}, G={fgColor.G}, B={fgColor.B}\n" +
                        $"Background Applied: {bgProperty != null && bgProperty.CanWrite}\n" +
                        $"Foreground Applied: {fgProperty != null && fgProperty.CanWrite}";
                    
                    MessageBox.Show(diagnosticMsg, "ThemeManager - ApplyTheme()", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"Failed to apply theme to {element.GetType().Name}");
                
                if (DiagnosticMode)
                {
                    MessageBox.Show($"Failed to apply theme: {ex.Message}", "ThemeManager - ApplyTheme() ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// Creates a SolidColorBrush for the background
        /// </summary>
        public static SolidColorBrush GetBackgroundBrush()
        {
            return new SolidColorBrush(GetBackgroundColor());
        }

        /// <summary>
        /// Creates a SolidColorBrush for the foreground
        /// </summary>
        public static SolidColorBrush GetForegroundBrush()
        {
            return new SolidColorBrush(GetForegroundColor());
        }

        /// <summary>
        /// Creates a SolidColorBrush for headers
        /// </summary>
        public static SolidColorBrush GetHeaderBackgroundBrush()
        {
            return new SolidColorBrush(GetHeaderBackgroundColor());
        }
        
        /// <summary>
        /// Shows a comprehensive diagnostic message about the current theme
        /// </summary>
        public static void ShowDiagnostics()
        {
            try
            {
                bool isDark = IsDarkTheme();
                var bgColor = GetBackgroundColor();
                var fgColor = GetForegroundColor();
                var headerColor = GetHeaderBackgroundColor();
                
                string message = $"=== SSMS Theme Diagnostics ===\n\n" +
                    $"Theme Type: {(isDark ? "DARK" : "LIGHT")}\n\n" +
                    $"Background Color:\n" +
                    $"  RGB: ({bgColor.R}, {bgColor.G}, {bgColor.B})\n" +
                    $"  Hex: #{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}\n\n" +
                    $"Foreground (Text) Color:\n" +
                    $"  RGB: ({fgColor.R}, {fgColor.G}, {fgColor.B})\n" +
                    $"  Hex: #{fgColor.R:X2}{fgColor.G:X2}{fgColor.B:X2}\n\n" +
                    $"Header Background Color:\n" +
                    $"  RGB: ({headerColor.R}, {headerColor.G}, {headerColor.B})\n" +
                    $"  Hex: #{headerColor.R:X2}{headerColor.G:X2}{headerColor.B:X2}\n\n" +
                    $"Diagnostic Mode: {DiagnosticMode}";
                
                Logger.Info(message);
                MessageBox.Show(message, "ThemeManager Diagnostics", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Failed to gather diagnostics:\n{ex.Message}\n\n{ex.StackTrace}";
                Logger.Error(ex, "ShowDiagnostics failed");
                MessageBox.Show(errorMsg, "ThemeManager Diagnostics ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
