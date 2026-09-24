namespace AxialSqlTools
{
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Reflection;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Controls.Primitives;
    using System.Windows.Media;
    using System.Windows.Navigation;
    using Microsoft.VisualStudio.PlatformUI;
    using Microsoft.VisualStudio.Shell;

    internal static class VsThemeBrushResolver
    {
        public static Brush ResolveBrush(FrameworkElement scope, object resourceKey)
        {
            if (resourceKey == null)
            {
                return null;
            }

            return scope?.TryFindResource(resourceKey) as Brush
                ?? Application.Current?.TryFindResource(resourceKey) as Brush;
        }

        public static Brush ResolveEnvironmentBrushByName(FrameworkElement scope, string keyName)
        {
            if (string.IsNullOrWhiteSpace(keyName))
            {
                return null;
            }

            PropertyInfo property = typeof(EnvironmentColors).GetProperty(keyName, BindingFlags.Public | BindingFlags.Static);
            object key = property?.GetValue(null);
            return ResolveBrush(scope, key);
        }

        public static Color GetBrushColor(Brush brush, Color fallback)
        {
            if (brush is SolidColorBrush solidBrush)
            {
                return solidBrush.Color;
            }

            return fallback;
        }

        public static double GetRelativeLuminance(Color color)
        {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;

            double rLinear = r <= 0.03928 ? r / 12.92 : Math.Pow((r + 0.055) / 1.055, 2.4);
            double gLinear = g <= 0.03928 ? g / 12.92 : Math.Pow((g + 0.055) / 1.055, 2.4);
            double bLinear = b <= 0.03928 ? b / 12.92 : Math.Pow((b + 0.055) / 1.055, 2.4);

            return 0.2126 * rLinear + 0.7152 * gLinear + 0.0722 * bLinear;
        }

        public static Color BlendColors(Color baseColor, Color blendColor, double blendAmount)
        {
            blendAmount = Math.Max(0.0, Math.Min(1.0, blendAmount));
            byte r = (byte)Math.Round((baseColor.R * (1.0 - blendAmount)) + (blendColor.R * blendAmount));
            byte g = (byte)Math.Round((baseColor.G * (1.0 - blendAmount)) + (blendColor.G * blendAmount));
            byte b = (byte)Math.Round((baseColor.B * (1.0 - blendAmount)) + (blendColor.B * blendAmount));
            return Color.FromRgb(r, g, b);
        }

        // Preserve a semantic color where possible, including on custom themes.
        public static Color EnsureTextContrast(Color color, Color background, Color fallback)
        {
            double backgroundLuminance = GetRelativeLuminance(background);
            for (int step = 0; step <= 20; step++)
            {
                Color candidate = BlendColors(color, fallback, step / 20.0);
                double luminance = GetRelativeLuminance(candidate);
                double contrast = (Math.Max(luminance, backgroundLuminance) + 0.05)
                    / (Math.Min(luminance, backgroundLuminance) + 0.05);
                if (contrast >= 4.5) return candidate;
            }
            return fallback;
        }
    }

    internal static class ToolWindowThemeResources
    {
        public static void ApplySharedTheme(FrameworkElement control)
        {
            Brush bg = VsThemeBrushResolver.ResolveBrush(control, EnvironmentColors.ToolWindowBackgroundBrushKey)
                ?? SystemColors.WindowBrush;
            Brush fg = VsThemeBrushResolver.ResolveBrush(control, EnvironmentColors.ToolWindowTextBrushKey)
                ?? SystemColors.WindowTextBrush;
            Brush border = VsThemeBrushResolver.ResolveBrush(control, EnvironmentColors.ToolWindowBorderBrushKey)
                ?? SystemColors.ActiveBorderBrush;
            Brush link = VsThemeBrushResolver.ResolveBrush(control, EnvironmentColors.ControlLinkTextBrushKey)
                ?? SystemColors.HotTrackBrush;
            Brush accent = VsThemeBrushResolver.ResolveEnvironmentBrushByName(control, "MainWindowActiveDefaultBorderBrushKey")
                ?? VsThemeBrushResolver.ResolveEnvironmentBrushByName(control, "SystemAccentBrushKey")
                ?? border;
            Brush success = VsThemeBrushResolver.ResolveEnvironmentBrushByName(control, "SystemGreenTextBrushKey")
                ?? new SolidColorBrush(Color.FromRgb(0x10, 0x7C, 0x10));
            Brush error = VsThemeBrushResolver.ResolveEnvironmentBrushByName(control, "SystemRedTextBrushKey")
                ?? new SolidColorBrush(Color.FromRgb(0xA1, 0x26, 0x0D));

            Color bgColor = VsThemeBrushResolver.GetBrushColor(bg, Colors.White);
            Color fgColor = VsThemeBrushResolver.GetBrushColor(fg, Colors.Black);
            success = new SolidColorBrush(VsThemeBrushResolver.EnsureTextContrast(
                VsThemeBrushResolver.GetBrushColor(success, fgColor), bgColor, fgColor));
            error = VsThemeBrushResolver.ResolveBrush(control, EnvironmentColors.ToolWindowValidationErrorTextBrushKey)
                ?? new SolidColorBrush(VsThemeBrushResolver.EnsureTextContrast(
                    VsThemeBrushResolver.GetBrushColor(error, fgColor), bgColor, fgColor));
            Color accentColor = VsThemeBrushResolver.GetBrushColor(accent, Color.FromRgb(0x00, 0x7A, 0xCC));
            Color successColor = VsThemeBrushResolver.GetBrushColor(success, Color.FromRgb(0x10, 0x7C, 0x10));
            Color errorColor = VsThemeBrushResolver.GetBrushColor(error, Color.FromRgb(0xA1, 0x26, 0x0D));
            bool isLightTheme = VsThemeBrushResolver.GetRelativeLuminance(bgColor) > 0.6;

            Color headerColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.04)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.04);
            Color subtleBorderColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.16)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.14);
            Color primaryHoverColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(accentColor, Colors.White, 0.10)
                : VsThemeBrushResolver.BlendColors(accentColor, Colors.White, 0.14);
            Color primaryPressedColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(accentColor, Colors.Black, 0.12)
                : VsThemeBrushResolver.BlendColors(accentColor, Colors.Black, 0.16);
            Color primaryForegroundColor = VsThemeBrushResolver.GetRelativeLuminance(accentColor) > 0.179
                ? Colors.Black
                : Colors.White;
            Color dangerHoverColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(errorColor, Colors.White, 0.08)
                : VsThemeBrushResolver.BlendColors(errorColor, Colors.White, 0.10);
            Color dangerPressedColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(errorColor, Colors.Black, 0.12)
                : VsThemeBrushResolver.BlendColors(errorColor, Colors.Black, 0.16);
            Color dangerForegroundColor = VsThemeBrushResolver.GetRelativeLuminance(errorColor) > 0.179
                ? Colors.Black
                : Colors.White;
            Color gridHeaderColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.03)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.08);
            Color gridAlternateRowColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.02)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.04);
            Color diffInsertedBackgroundColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, successColor, 0.20)
                : VsThemeBrushResolver.BlendColors(bgColor, successColor, 0.36);
            Color diffDeletedBackgroundColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, errorColor, 0.20)
                : VsThemeBrushResolver.BlendColors(bgColor, errorColor, 0.36);
            Color diffModifiedBackgroundColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, fgColor, 0.10)
                : VsThemeBrushResolver.BlendColors(bgColor, fgColor, 0.20);
            Color diffInsertedForegroundColor = VsThemeBrushResolver.GetRelativeLuminance(diffInsertedBackgroundColor) > 0.179
                ? Colors.Black
                : Colors.White;
            Color diffDeletedForegroundColor = VsThemeBrushResolver.GetRelativeLuminance(diffDeletedBackgroundColor) > 0.179
                ? Colors.Black
                : Colors.White;
            Color diffModifiedForegroundColor = VsThemeBrushResolver.GetRelativeLuminance(diffModifiedBackgroundColor) > 0.179
                ? Colors.Black
                : Colors.White;

            control.Resources["AxialThemeBackgroundBrush"] = bg;
            control.Resources["AxialThemeForegroundBrush"] = fg;
            control.Resources["AxialThemeBorderBrush"] = border;
            control.Resources["AxialThemeSubtleBorderBrush"] = new SolidColorBrush(subtleBorderColor);
            control.Resources["AxialThemeHeaderBackgroundBrush"] = new SolidColorBrush(headerColor);
            control.Resources["AxialThemeLinkBrush"] = link;
            control.Resources["AxialThemeAccentBrush"] = new SolidColorBrush(accentColor);
            control.Resources["AxialThemeStatusErrorBrush"] = error;
            control.Resources["AxialThemeStatusSuccessBrush"] = success;
            control.Resources["AxialThemePrimaryButtonBackgroundBrush"] = new SolidColorBrush(accentColor);
            control.Resources["AxialThemePrimaryButtonHoverBrush"] = new SolidColorBrush(primaryHoverColor);
            control.Resources["AxialThemePrimaryButtonPressedBrush"] = new SolidColorBrush(primaryPressedColor);
            control.Resources["AxialThemePrimaryButtonForegroundBrush"] = new SolidColorBrush(primaryForegroundColor);
            control.Resources["AxialThemeDangerButtonBackgroundBrush"] = new SolidColorBrush(errorColor);
            control.Resources["AxialThemeDangerButtonHoverBrush"] = new SolidColorBrush(dangerHoverColor);
            control.Resources["AxialThemeDangerButtonPressedBrush"] = new SolidColorBrush(dangerPressedColor);
            control.Resources["AxialThemeDangerButtonForegroundBrush"] = new SolidColorBrush(dangerForegroundColor);
            control.Resources["AxialThemeGridHeaderBackgroundBrush"] = new SolidColorBrush(gridHeaderColor);
            control.Resources["AxialThemeGridAlternateRowBrush"] = new SolidColorBrush(gridAlternateRowColor);
            control.Resources["AxialThemeDiffInsertedBackgroundBrush"] = new SolidColorBrush(diffInsertedBackgroundColor);
            control.Resources["AxialThemeDiffInsertedForegroundBrush"] = new SolidColorBrush(diffInsertedForegroundColor);
            control.Resources["AxialThemeDiffDeletedBackgroundBrush"] = new SolidColorBrush(diffDeletedBackgroundColor);
            control.Resources["AxialThemeDiffDeletedForegroundBrush"] = new SolidColorBrush(diffDeletedForegroundColor);
            control.Resources["AxialThemeDiffModifiedBackgroundBrush"] = new SolidColorBrush(diffModifiedBackgroundColor);
            control.Resources["AxialThemeDiffModifiedForegroundBrush"] = new SolidColorBrush(diffModifiedForegroundColor);

            // Keep foreground/background pairs from the host together. In particular,
            // selected text must not inherit the ordinary tool-window foreground.
            SetBrush(control, "AxialThemeGridSelectionBrush", EnvironmentColors.SystemHighlightBrushKey, SystemColors.HighlightBrush);
            SetBrush(control, "AxialThemeGridSelectionTextBrush", EnvironmentColors.SystemHighlightTextBrushKey, SystemColors.HighlightTextBrush);
            SetBrush(control, "AxialThemeDisabledForegroundBrush", EnvironmentColors.SystemGrayTextBrushKey, SystemColors.GrayTextBrush);
            SetBrush(control, "AxialThemeInputBackgroundBrush", CommonControlsColors.TextBoxBackgroundBrushKey, bg);
            SetBrush(control, "AxialThemeInputForegroundBrush", CommonControlsColors.TextBoxTextBrushKey, fg);
            SetBrush(control, "AxialThemeButtonBackgroundBrush", CommonControlsColors.ButtonBrushKey, bg);
            SetBrush(control, "AxialThemeButtonForegroundBrush", CommonControlsColors.ButtonTextBrushKey, fg);
            SetBrush(control, "AxialThemeButtonHoverBrush", CommonControlsColors.ButtonHoverBrushKey, bg);
            SetBrush(control, "AxialThemeButtonHoverTextBrush", CommonControlsColors.ButtonHoverTextBrushKey, fg);
            SetBrush(control, "AxialThemeButtonPressedBrush", CommonControlsColors.ButtonPressedBrushKey, bg);
            SetBrush(control, "AxialThemeButtonPressedTextBrush", CommonControlsColors.ButtonPressedTextBrushKey, fg);
            SetBrush(control, "AxialThemeTabHeaderBackgroundBrush", CommonControlsColors.InnerTabInactiveBackgroundBrushKey, bg);
            SetBrush(control, "AxialThemeTabHeaderTextBrush", CommonControlsColors.InnerTabInactiveTextBrushKey, fg);
            SetBrush(control, "AxialThemeTabHeaderHoverBrush", CommonControlsColors.InnerTabInactiveHoverBackgroundBrushKey, bg);
            SetBrush(control, "AxialThemeTabHeaderHoverTextBrush", CommonControlsColors.InnerTabInactiveHoverTextBrushKey, fg);
            SetBrush(control, "AxialThemeTabHeaderSelectedBrush", CommonControlsColors.InnerTabActiveBackgroundBrushKey, bg);
            SetBrush(control, "AxialThemeTabHeaderSelectedTextBrush", CommonControlsColors.InnerTabActiveTextBrushKey, fg);
            SetBrush(control, "AxialThemeToolTipBackgroundBrush", EnvironmentColors.ToolTipBrushKey, bg);
            SetBrush(control, "AxialThemeToolTipForegroundBrush", EnvironmentColors.ToolTipTextBrushKey, fg);

            if (SystemParameters.HighContrast) ApplyHighContrast(control);
            ApplyHostControlStyles(control);
        }

        private static void SetBrush(FrameworkElement scope, string name, object hostKey, Brush fallback)
        {
            scope.Resources[name] = VsThemeBrushResolver.ResolveBrush(scope, hostKey) ?? fallback;
        }

        private static void ApplyHighContrast(FrameworkElement scope)
        {
            // Do not blend, dim or invent semantic colors in a high-contrast theme.
            foreach (string name in new[] { "Background", "HeaderBackground", "InputBackground", "GridHeaderBackground",
                "GridAlternateRow", "TabHeaderBackground", "TabHeaderHover", "ButtonBackground", "ButtonHover", "ButtonPressed",
                "DiffInsertedBackground", "DiffDeletedBackground", "DiffModifiedBackground", "ToolTipBackground" })
                scope.Resources["AxialTheme" + name + "Brush"] = SystemColors.WindowBrush;
            foreach (string name in new[] { "Foreground", "InputForeground", "ButtonForeground", "ButtonHoverText", "ButtonPressedText",
                "TabHeaderText", "TabHeaderHoverText", "Border", "SubtleBorder", "Accent", "StatusError", "StatusSuccess",
                "DiffInsertedForeground", "DiffDeletedForeground", "DiffModifiedForeground", "ToolTipForeground" })
                scope.Resources["AxialTheme" + name + "Brush"] = SystemColors.WindowTextBrush;
            foreach (string name in new[] { "GridSelection", "TabHeaderSelected", "PrimaryButtonBackground", "PrimaryButtonHover",
                "PrimaryButtonPressed", "DangerButtonBackground", "DangerButtonHover", "DangerButtonPressed" })
                scope.Resources["AxialTheme" + name + "Brush"] = SystemColors.HighlightBrush;
            foreach (string name in new[] { "GridSelectionText", "TabHeaderSelectedText", "PrimaryButtonForeground", "DangerButtonForeground" })
                scope.Resources["AxialTheme" + name + "Brush"] = SystemColors.HighlightTextBrush;
            scope.Resources["AxialThemeLinkBrush"] = SystemColors.HotTrackBrush;
            scope.Resources["AxialThemeDisabledForegroundBrush"] = SystemColors.GrayTextBrush;
        }

        private static void ApplyHostControlStyles(FrameworkElement scope)
        {
            // The shell templates cover focus, disabled, hover, popup and editable
            // states that cannot be themed by setting Background/Foreground alone.
            // Keep shared XAML styles as a fallback for hosts missing a resource.
            UseHostStyle(scope, typeof(Button), VsResourceKeys.ThemedDialogButtonStyleKey);
            UseHostStyle(scope, typeof(CheckBox), VsResourceKeys.ThemedDialogCheckBoxStyleKey);
            UseHostStyle(scope, typeof(RadioButton), VsResourceKeys.ThemedDialogRadioButtonStyleKey);
            UseHostStyle(scope, typeof(ToggleButton), VsResourceKeys.ThemedDialogToggleButtonStyleKey);
            UseHostStyle(scope, typeof(TextBox), VsResourceKeys.ThemedDialogTextBoxStyleKey);
            UseHostStyle(scope, typeof(ComboBox), VsResourceKeys.ThemedDialogComboBoxStyleKey);
            UseHostStyle(scope, typeof(ComboBoxItem), VsResourceKeys.ComboBoxItemStyleKey);
            UseHostStyle(scope, typeof(ListBox), VsResourceKeys.ThemedDialogListBoxStyleKey);
            UseHostStyle(scope, typeof(ListView), VsResourceKeys.ThemedDialogListViewStyleKey);
            UseHostStyle(scope, typeof(ListViewItem), VsResourceKeys.ThemedDialogListViewItemStyleKey);
            UseHostStyle(scope, typeof(GridViewColumnHeader), VsResourceKeys.ThemedDialogGridViewColumnHeaderStyleKey);
            UseHostStyle(scope, typeof(TreeView), VsResourceKeys.ThemedDialogTreeViewStyleKey);
            UseHostStyle(scope, typeof(TreeViewItem), VsResourceKeys.ThemedDialogTreeViewItemStyleKey);
            UseHostStyle(scope, typeof(ScrollBar), VsResourceKeys.ScrollBarStyleKey);
            UseHostStyle(scope, typeof(ScrollViewer), VsResourceKeys.ScrollViewerStyleKey);
            UseHostStyle(scope, typeof(ProgressBar), VsResourceKeys.ProgressBarStyleKey);
        }

        private static void UseHostStyle(FrameworkElement scope, Type targetType, object key)
        {
            var defaults = scope.TryFindResource(VsResourceKeys.ThemedDialogDefaultStylesKey) as ResourceDictionary;
            var style = scope.TryFindResource(key) as Style ?? defaults?[targetType] as Style;
            if (style != null && style.TargetType.IsAssignableFrom(targetType) && !ReferenceEquals(scope.Resources[targetType], style))
                scope.Resources[targetType] = style;
        }
    }

    internal sealed class ToolWindowThemeController : IDisposable
    {
        private readonly FrameworkElement control;
        private readonly Action applyTheme;
        private bool isThemeSubscribed;
        private bool disposed;

        public ToolWindowThemeController(FrameworkElement control, Action applyTheme)
        {
            this.control = control ?? throw new ArgumentNullException(nameof(control));
            this.applyTheme = applyTheme ?? throw new ArgumentNullException(nameof(applyTheme));

            this.control.Loaded += OnLoaded;
            this.control.Unloaded += OnUnloaded;
            this.control.IsVisibleChanged += OnIsVisibleChanged;
            if (control is Window window) window.Closed += OnClosed;

            this.applyTheme();
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (disposed) return;
            applyTheme();
            SubscribeToThemeChanges();
        }

        private void OnUnloaded(object sender, RoutedEventArgs e)
        {
            UnsubscribeFromThemeChanges();
        }

        private void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!disposed && control.IsVisible)
            {
                applyTheme();
            }
        }

        private void SubscribeToThemeChanges()
        {
            if (isThemeSubscribed)
            {
                return;
            }

            VSColorTheme.ThemeChanged += OnVsThemeChanged;
            SystemParameters.StaticPropertyChanged += OnSystemParametersChanged;
            isThemeSubscribed = true;
        }

        private void UnsubscribeFromThemeChanges()
        {
            if (!isThemeSubscribed)
            {
                return;
            }

            VSColorTheme.ThemeChanged -= OnVsThemeChanged;
            SystemParameters.StaticPropertyChanged -= OnSystemParametersChanged;
            isThemeSubscribed = false;
        }

        private void OnVsThemeChanged(ThemeChangedEventArgs e)
        {
            QueueThemeRefresh();
        }

        private void OnSystemParametersChanged(object sender, PropertyChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(e.PropertyName) || e.PropertyName == nameof(SystemParameters.HighContrast)) QueueThemeRefresh();
        }

        private void QueueThemeRefresh()
        {
            if (disposed || control.Dispatcher.HasShutdownStarted) return;
            // Wait until the shell has replaced its WPF resources. Ignore queued
            // work after unloading or closing the window.
            _ = control.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (!disposed && isThemeSubscribed) applyTheme();
            }), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private void OnClosed(object sender, EventArgs e) => Dispose();

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            control.Loaded -= OnLoaded;
            control.Unloaded -= OnUnloaded;
            control.IsVisibleChanged -= OnIsVisibleChanged;
            if (control is Window window) window.Closed -= OnClosed;
            UnsubscribeFromThemeChanges();
        }
    }

    internal static class ToolWindowNavigation
    {
        public static bool OpenExternalUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            Process.Start(new ProcessStartInfo(url)
            {
                UseShellExecute = true,
            });

            return true;
        }

        public static void HandleRequestNavigate(RequestNavigateEventArgs e)
        {
            if (e == null)
            {
                return;
            }

            if (OpenExternalUrl(e.Uri?.AbsoluteUri))
            {
                e.Handled = true;
            }
        }
    }
}
