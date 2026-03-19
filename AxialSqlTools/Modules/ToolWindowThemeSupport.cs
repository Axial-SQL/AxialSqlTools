namespace AxialSqlTools
{
    using System;
    using System.Windows;
    using System.Windows.Media;
    using Microsoft.VisualStudio.PlatformUI;
    using Microsoft.VisualStudio.Shell;

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
            Color accentColor = VsThemeBrushResolver.GetBrushColor(accent, Color.FromRgb(0x00, 0x7A, 0xCC));
            Color errorColor = VsThemeBrushResolver.GetBrushColor(error, Color.FromRgb(0xA1, 0x26, 0x0D));
            bool isLightTheme = VsThemeBrushResolver.GetRelativeLuminance(bgColor) > 0.6;

            Color headerColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.04)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.04);
            Color tabHeaderColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.05)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.06);
            Color tabHoverColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.10)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.12);
            Color tabSelectedColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.16)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.18);
            Color buttonBackgroundColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.06)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.08);
            Color buttonHoverColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.14)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.10);
            Color buttonPressedColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.22)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.18);
            Color subtleBorderColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.16)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.14);
            Color primaryHoverColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(accentColor, Colors.White, 0.10)
                : VsThemeBrushResolver.BlendColors(accentColor, Colors.White, 0.14);
            Color primaryPressedColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(accentColor, Colors.Black, 0.12)
                : VsThemeBrushResolver.BlendColors(accentColor, Colors.Black, 0.16);
            Color primaryForegroundColor = VsThemeBrushResolver.GetRelativeLuminance(accentColor) > 0.55
                ? Colors.Black
                : Colors.White;
            Color dangerHoverColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(errorColor, Colors.White, 0.08)
                : VsThemeBrushResolver.BlendColors(errorColor, Colors.White, 0.10);
            Color dangerPressedColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(errorColor, Colors.Black, 0.12)
                : VsThemeBrushResolver.BlendColors(errorColor, Colors.Black, 0.16);
            Color dangerForegroundColor = VsThemeBrushResolver.GetRelativeLuminance(errorColor) > 0.55
                ? Colors.Black
                : Colors.White;
            Color gridHeaderColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.03)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.08);
            Color gridAlternateRowColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.02)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.04);
            Color gridSelectionColor = isLightTheme
                ? VsThemeBrushResolver.BlendColors(bgColor, Colors.Black, 0.12)
                : VsThemeBrushResolver.BlendColors(bgColor, Colors.White, 0.14);

            control.Resources["AxialThemeBackgroundBrush"] = bg;
            control.Resources["AxialThemeForegroundBrush"] = fg;
            control.Resources["AxialThemeBorderBrush"] = border;
            control.Resources["AxialThemeSubtleBorderBrush"] = new SolidColorBrush(subtleBorderColor);
            control.Resources["AxialThemeHeaderBackgroundBrush"] = new SolidColorBrush(headerColor);
            control.Resources["AxialThemeLinkBrush"] = link;
            control.Resources["AxialThemeAccentBrush"] = new SolidColorBrush(accentColor);
            control.Resources["AxialThemeStatusErrorBrush"] = error;
            control.Resources["AxialThemeStatusSuccessBrush"] = success;
            control.Resources["AxialThemeTabHeaderBackgroundBrush"] = new SolidColorBrush(tabHeaderColor);
            control.Resources["AxialThemeTabHeaderHoverBrush"] = new SolidColorBrush(tabHoverColor);
            control.Resources["AxialThemeTabHeaderSelectedBrush"] = new SolidColorBrush(tabSelectedColor);
            control.Resources["AxialThemeButtonBackgroundBrush"] = new SolidColorBrush(buttonBackgroundColor);
            control.Resources["AxialThemeButtonHoverBrush"] = new SolidColorBrush(buttonHoverColor);
            control.Resources["AxialThemeButtonPressedBrush"] = new SolidColorBrush(buttonPressedColor);
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
            control.Resources["AxialThemeGridSelectionBrush"] = new SolidColorBrush(gridSelectionColor);
        }
    }

    internal sealed class ToolWindowThemeController : IDisposable
    {
        private readonly FrameworkElement control;
        private readonly Action applyTheme;
        private bool isThemeSubscribed;

        public ToolWindowThemeController(FrameworkElement control, Action applyTheme)
        {
            this.control = control ?? throw new ArgumentNullException(nameof(control));
            this.applyTheme = applyTheme ?? throw new ArgumentNullException(nameof(applyTheme));

            this.control.Loaded += OnLoaded;
            this.control.Unloaded += OnUnloaded;
            this.control.IsVisibleChanged += OnIsVisibleChanged;

            this.applyTheme();
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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

            if (control.IsVisible)
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
            isThemeSubscribed = true;
        }

        private void UnsubscribeFromThemeChanges()
        {
            if (!isThemeSubscribed)
            {
                return;
            }

            VSColorTheme.ThemeChanged -= OnVsThemeChanged;
            isThemeSubscribed = false;
        }

        private void OnVsThemeChanged(ThemeChangedEventArgs e)
        {
            if (!control.Dispatcher.CheckAccess())
            {
                control.Dispatcher.BeginInvoke(new Action(applyTheme));
                return;
            }

            applyTheme();
        }

        public void Dispose()
        {
            control.Loaded -= OnLoaded;
            control.Unloaded -= OnUnloaded;
            control.IsVisibleChanged -= OnIsVisibleChanged;
            UnsubscribeFromThemeChanges();
        }
    }
}
