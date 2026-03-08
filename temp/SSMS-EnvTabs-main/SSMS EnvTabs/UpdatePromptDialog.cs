using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Reflection;
using Microsoft.VisualStudio.PlatformUI;
using WinFormsDialogResult = System.Windows.Forms.DialogResult;

namespace SSMS_EnvTabs
{
    internal sealed class UpdatePromptDialog : DialogWindow, IDisposable
    {
        private readonly Action releaseNotesRequested;
        private readonly Action openConfigRequested;
        private WinFormsDialogResult dialogResult = WinFormsDialogResult.Cancel;
        private Button installButton;
        private Button openConfigButton;
        private Button laterButton;

        public UpdatePromptDialog(string latestVersion, string currentVersion, Action releaseNotesRequested, Action openConfigRequested)
        {
            this.releaseNotesRequested = releaseNotesRequested;
            this.openConfigRequested = openConfigRequested;

            Title = "SSMS EnvTabs Update";
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            SizeToContent = SizeToContent.Height;
            Width = 480;
            MinWidth = 480;

            BuildUi(latestVersion, currentVersion);
            ApplyThemeResources();
        }

        public new WinFormsDialogResult ShowDialog()
        {
            base.ShowDialog();
            return dialogResult;
        }

        public void Dispose()
        {
            if (IsVisible)
            {
                Close();
            }
        }

        private void BuildUi(string latestVersion, string currentVersion)
        {
            var root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var contentPanel = new StackPanel
            {
                Margin = new Thickness(16, 14, 16, 10),
                Orientation = Orientation.Vertical
            };

            var header = new TextBlock
            {
                Text = "EnvTabs update available!",
                FontWeight = FontWeights.Bold,
                FontSize = SystemFonts.MessageFontSize + 6,
                Margin = new Thickness(0, 0, 0, 10)
            };
            contentPanel.Children.Add(header);

            var versionGrid = new Grid
            {
                Margin = new Thickness(0, 0, 0, 10)
            };
            versionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            versionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            versionGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            versionGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            AddVersionRow(versionGrid, 0, "Current:", currentVersion);
            AddVersionRow(versionGrid, 1, "Available:", latestVersion);

            contentPanel.Children.Add(versionGrid);

            var releaseNotes = new TextBlock
            {
                Margin = new Thickness(0, 8, 0, 6)
            };
            var link = new Hyperlink(new Run("Full release notes here."));
            link.SetResourceReference(TextElement.ForegroundProperty, EnvironmentColors.ControlLinkTextBrushKey);
            link.Click += (s, e) => releaseNotesRequested?.Invoke();
            releaseNotes.Inlines.Add(link);
            contentPanel.Children.Add(releaseNotes);

            var configNote = new TextBlock
            {
                Text = "To disable update checking, press \"Open Config\" below.",
                TextWrapping = TextWrapping.Wrap
            };
            contentPanel.Children.Add(configNote);

            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 6, 0, 8)
            };

            installButton = new Button
            {
                Content = "_Install",
                MinWidth = 110,
                Height = 26,
                Margin = new Thickness(6, 10, 6, 10),
                IsDefault = true
            };
            ApplyControlBrushes(installButton);
            installButton.Click += (s, e) =>
            {
                dialogResult = WinFormsDialogResult.Yes;
                this.DialogResult = true;
            };

            openConfigButton = new Button
            {
                Content = "_Open Config",
                MinWidth = 120,
                Height = 26,
                Margin = new Thickness(6, 10, 6, 10)
            };
            ApplyControlBrushes(openConfigButton);
            openConfigButton.Click += (s, e) =>
            {
                openConfigRequested?.Invoke();
                dialogResult = WinFormsDialogResult.OK;
                this.DialogResult = true;
            };

            laterButton = new Button
            {
                Content = "_Later",
                MinWidth = 110,
                Height = 26,
                Margin = new Thickness(6, 10, 6, 10),
                IsCancel = true
            };
            ApplyControlBrushes(laterButton);
            laterButton.Click += (s, e) =>
            {
                dialogResult = WinFormsDialogResult.Cancel;
                this.DialogResult = false;
            };

            buttonPanel.Children.Add(installButton);
            buttonPanel.Children.Add(openConfigButton);
            buttonPanel.Children.Add(laterButton);

            Grid.SetRow(contentPanel, 0);
            Grid.SetRow(buttonPanel, 1);

            root.Children.Add(contentPanel);
            root.Children.Add(buttonPanel);

            Content = root;
        }

        private void ApplyThemeResources()
        {
            SetResourceReference(BackgroundProperty, EnvironmentColors.ToolWindowBackgroundBrushKey);
            SetResourceReference(ForegroundProperty, EnvironmentColors.ToolWindowTextBrushKey);

            if (Content is FrameworkElement root)
            {
                root.SetResourceReference(BackgroundProperty, EnvironmentColors.ToolWindowBackgroundBrushKey);
                root.SetResourceReference(TextElement.ForegroundProperty, EnvironmentColors.ToolWindowTextBrushKey);
            }

            ApplyDialogButtonStyles();
        }

        private void ApplyDialogButtonStyles()
        {
            var baseBackground = TryFindResource(EnvironmentColors.ToolWindowBackgroundBrushKey) as Brush
                ?? TryFindResource(SystemColors.ControlBrushKey) as Brush
                ?? Brushes.Transparent;

            var baseForeground = TryFindResource(EnvironmentColors.ToolWindowTextBrushKey) as Brush
                ?? TryFindResource(SystemColors.ControlTextBrushKey) as Brush
                ?? Brushes.Black;

            bool isLightTheme = GetRelativeLuminance(GetBrushColor(baseBackground, Colors.White)) > 0.6;

            var primaryBaseColor = (Color)ColorConverter.ConvertFromString(isLightTheme ? "#5649B0" : "#9184EE");
            var primaryHoverColor = (Color)ColorConverter.ConvertFromString(isLightTheme ? "#665bb7" : "#867bda");

            var accentBrush = new SolidColorBrush(primaryBaseColor);
            var primaryHover = new SolidColorBrush(primaryHoverColor);
            var primaryPressed = BlendBrush(accentBrush, Colors.Black, 0.12f);

            var accentForeground = isLightTheme ? Brushes.White : Brushes.Black;
            var secondaryForeground = isLightTheme ? Brushes.Black : Brushes.White;

            var baseBorder = WithOpacity(secondaryForeground, 0.28);
            var secondaryHover = isLightTheme
                ? BlendBrush(baseBackground, Colors.Black, 0.06f)
                : BlendBrush(baseBackground, Colors.White, 0.06f);
            var secondaryPressed = isLightTheme
                ? BlendBrush(baseBackground, Colors.Black, 0.12f)
                : BlendBrush(baseBackground, Colors.Black, 0.08f);

            var primaryStyle = CreateButtonStyle(accentBrush, baseBorder, accentForeground, primaryHover, primaryPressed);
            var secondaryStyle = CreateButtonStyle(baseBackground, baseBorder, secondaryForeground, secondaryHover, secondaryPressed);

            ApplyButtonStyle(installButton, primaryStyle);
            ApplyButtonStyle(openConfigButton, secondaryStyle);
            ApplyButtonStyle(laterButton, secondaryStyle);
        }

        private static Color GetBrushColor(Brush brush, Color fallback)
        {
            if (brush is SolidColorBrush solidBrush)
            {
                return solidBrush.Color;
            }

            return fallback;
        }

        private static double GetRelativeLuminance(Color color)
        {
            return (0.2126 * color.R + 0.7152 * color.G + 0.0722 * color.B) / 255.0;
        }

        private static void ApplyButtonStyle(Button button, Style style)
        {
            if (button == null || style == null)
            {
                return;
            }

            button.ClearValue(BackgroundProperty);
            button.ClearValue(ForegroundProperty);
            button.ClearValue(BorderBrushProperty);

            button.Style = style;
            button.MinHeight = 24;
            button.Padding = new Thickness(12, 4, 12, 4);
        }

        private Style CreateButtonStyle(Brush background, Brush border, Brush foreground, Brush hoverBackground, Brush pressedBackground)
        {
            var style = new Style(typeof(Button));
            style.Setters.Add(new Setter(BackgroundProperty, background));
            style.Setters.Add(new Setter(BorderBrushProperty, border));
            style.Setters.Add(new Setter(ForegroundProperty, foreground));
            style.Setters.Add(new Setter(BorderThicknessProperty, new Thickness(1)));
            style.Setters.Add(new Setter(HorizontalContentAlignmentProperty, HorizontalAlignment.Center));
            style.Setters.Add(new Setter(VerticalContentAlignmentProperty, VerticalAlignment.Center));

            var template = new ControlTemplate(typeof(Button));
            var borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.Name = "Bd";
            borderFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(BackgroundProperty));
            borderFactory.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(BorderBrushProperty));
            borderFactory.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(BorderThicknessProperty));
            borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(3));

            var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
            contentPresenter.SetValue(ContentPresenter.ContentProperty, new TemplateBindingExtension(ContentControl.ContentProperty));
            contentPresenter.SetValue(ContentPresenter.ContentTemplateProperty, new TemplateBindingExtension(ContentControl.ContentTemplateProperty));
            contentPresenter.SetValue(ContentPresenter.ContentStringFormatProperty, new TemplateBindingExtension(ContentControl.ContentStringFormatProperty));
            contentPresenter.SetValue(ContentPresenter.MarginProperty, new TemplateBindingExtension(Control.PaddingProperty));
            contentPresenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.RecognizesAccessKeyProperty, true);

            borderFactory.AppendChild(contentPresenter);
            template.VisualTree = borderFactory;

            var hoverTrigger = new Trigger { Property = IsMouseOverProperty, Value = true };
            hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, hoverBackground, "Bd"));

            var pressedTrigger = new Trigger { Property = Button.IsPressedProperty, Value = true };
            pressedTrigger.Setters.Add(new Setter(Border.BackgroundProperty, pressedBackground, "Bd"));

            var disabledTrigger = new Trigger { Property = IsEnabledProperty, Value = false };
            disabledTrigger.Setters.Add(new Setter(OpacityProperty, 0.55));

            template.Triggers.Add(hoverTrigger);
            template.Triggers.Add(pressedTrigger);
            template.Triggers.Add(disabledTrigger);

            style.Setters.Add(new Setter(TemplateProperty, template));

            return style;
        }

        private Brush GetAccentBrush()
        {
            return TryGetEnvironmentBrush("AccentMediumBrushKey")
                ?? TryGetEnvironmentBrush("AccentDarkBrushKey")
                ?? TryGetEnvironmentBrush("AccentLightBrushKey")
                ?? TryFindResource(EnvironmentColors.ControlLinkTextBrushKey) as Brush
                ?? TryFindResource(SystemColors.HighlightBrushKey) as Brush;
        }

        private Brush TryGetEnvironmentBrush(string keyName)
        {
            var property = typeof(EnvironmentColors).GetProperty(keyName, BindingFlags.Public | BindingFlags.Static);
            if (property?.GetValue(null) is object resourceKey)
            {
                return TryFindResource(resourceKey) as Brush;
            }

            return null;
        }

        private static Brush WithOpacity(Brush brush, double opacity)
        {
            if (brush is SolidColorBrush solid)
            {
                var color = solid.Color;
                var alpha = (byte)Math.Max(0, Math.Min(255, opacity * 255));
                return new SolidColorBrush(Color.FromArgb(alpha, color.R, color.G, color.B));
            }

            return brush;
        }

        private static Brush BlendBrush(Brush brush, Color blendColor, float amount)
        {
            if (brush is SolidColorBrush solid)
            {
                var blended = BlendColor(solid.Color, blendColor, amount);
                return new SolidColorBrush(blended);
            }

            return brush;
        }

        private static Color BlendColor(Color baseColor, Color blendColor, float blendAmount)
        {
            blendAmount = Math.Max(0, Math.Min(1, blendAmount));
            byte r = (byte)(baseColor.R + (blendColor.R - baseColor.R) * blendAmount);
            byte g = (byte)(baseColor.G + (blendColor.G - baseColor.G) * blendAmount);
            byte b = (byte)(baseColor.B + (blendColor.B - baseColor.B) * blendAmount);
            return Color.FromArgb(baseColor.A, r, g, b);
        }

        private static void ApplyControlBrushes(Control control)
        {
            if (control == null)
            {
                return;
            }

            control.SetResourceReference(BackgroundProperty, EnvironmentColors.ToolWindowBackgroundBrushKey);
            control.SetResourceReference(ForegroundProperty, EnvironmentColors.ToolWindowTextBrushKey);
        }

        private static void AddVersionRow(Grid grid, int rowIndex, string label, string value)
        {
            var labelBlock = new TextBlock
            {
                Text = label,
                Margin = new Thickness(0, 2, 10, 2)
            };

            var valueBlock = new TextBlock
            {
                Text = value,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 2, 0, 2)
            };

            Grid.SetRow(labelBlock, rowIndex);
            Grid.SetColumn(labelBlock, 0);
            Grid.SetRow(valueBlock, rowIndex);
            Grid.SetColumn(valueBlock, 1);

            grid.Children.Add(labelBlock);
            grid.Children.Add(valueBlock);
        }

    }
}
