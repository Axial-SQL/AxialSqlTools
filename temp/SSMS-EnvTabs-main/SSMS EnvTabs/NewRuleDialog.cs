using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Interop;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Input;
using Microsoft.VisualStudio.PlatformUI;
using WinFormsDialogResult = System.Windows.Forms.DialogResult;

namespace SSMS_EnvTabs
{
    public sealed class NewRuleDialog : DialogWindow, IDisposable
    {
        public string RuleName { get; private set; }
        public int? SelectedColorIndex { get; private set; }
        public bool OpenConfigRequested { get; private set; }
        public string ServerAlias { get; private set; }
        public event Action<string> AliasConfirmed;

        private TextBox txtName;
        private ComboBox cmbColor;
        private Button btnSave;
        private Button btnCancel;
        private Button btnOpenConfig;

        private Grid panelAlias;
        private Grid panelRule;
        private TextBox txtAlias;
        private Button btnNext;
        private Button btnCancelAlias;

        private StackPanel panelRuleButtons;
        private StackPanel panelAliasButtons;

        private TextBlock lblHeader;
        private TextBlock lblServerValue;
        private TextBlock lblDatabaseValue;
        private TextBlock lblAliasHeader;

        private readonly string serverName;
        private readonly string databaseName;
        private readonly string existingAlias;
        private readonly string suggestedName;
        private readonly string suggestedGroupNameStyle;
        private readonly bool hideAliasStep;
        private readonly bool hideGroupNameRow;
        private readonly bool isEditMode;
        private readonly HashSet<int> usedColorIndexes;

        private WinFormsDialogResult dialogResult = WinFormsDialogResult.Cancel;

        private sealed class ColorItem
        {
            public int? Index { get; set; }
            public string Name { get; set; }
            public Color SwatchColor { get; set; }
            public string DisplayName { get; set; }
            public SolidColorBrush SwatchBrush => new SolidColorBrush(SwatchColor);
        }

        private static readonly List<ColorItem> ColorList = new List<ColorItem>
        {
            new ColorItem { Index = null, Name = "None", SwatchColor = Colors.Transparent },
            new ColorItem { Index = 0, Name = "Lavender", SwatchColor = (Color)ColorConverter.ConvertFromString("#9083ef") },
            new ColorItem { Index = 1, Name = "Gold", SwatchColor = (Color)ColorConverter.ConvertFromString("#d0b132") },
            new ColorItem { Index = 2, Name = "Cyan", SwatchColor = (Color)ColorConverter.ConvertFromString("#30b1cd") },
            new ColorItem { Index = 3, Name = "Burgundy", SwatchColor = (Color)ColorConverter.ConvertFromString("#cf6468") },
            new ColorItem { Index = 4, Name = "Green", SwatchColor = (Color)ColorConverter.ConvertFromString("#6ba12a") },
            new ColorItem { Index = 5, Name = "Brown", SwatchColor = (Color)ColorConverter.ConvertFromString("#bc8f6f") },
            new ColorItem { Index = 6, Name = "Royal Blue", SwatchColor = (Color)ColorConverter.ConvertFromString("#5bb2fa") },
            new ColorItem { Index = 7, Name = "Pumpkin", SwatchColor = (Color)ColorConverter.ConvertFromString("#d67441") },
            new ColorItem { Index = 8, Name = "Gray", SwatchColor = (Color)ColorConverter.ConvertFromString("#bdbcbc") },
            new ColorItem { Index = 9, Name = "Volt", SwatchColor = (Color)ColorConverter.ConvertFromString("#cbcc38") },
            new ColorItem { Index = 10, Name = "Teal", SwatchColor = (Color)ColorConverter.ConvertFromString("#2aa0a4") },
            new ColorItem { Index = 11, Name = "Magenta", SwatchColor = (Color)ColorConverter.ConvertFromString("#d957a7") },
            new ColorItem { Index = 12, Name = "Mint", SwatchColor = (Color)ColorConverter.ConvertFromString("#6bc6a5") },
            new ColorItem { Index = 13, Name = "Dark Brown", SwatchColor = (Color)ColorConverter.ConvertFromString("#946a5b") },
            new ColorItem { Index = 14, Name = "Blue", SwatchColor = (Color)ColorConverter.ConvertFromString("#6a8ec6") },
            new ColorItem { Index = 15, Name = "Pink", SwatchColor = (Color)ColorConverter.ConvertFromString("#e0a3a5") }
        };

        public sealed class NewRuleDialogOptions
        {
            public string Server { get; set; }
            public string Database { get; set; }
            public string SuggestedName { get; set; }
            public string SuggestedGroupNameStyle { get; set; }
            public int? SuggestedColorIndex { get; set; }
            public string ExistingAlias { get; set; }
            public bool HideDatabaseRow { get; set; }
            public bool HideAliasStep { get; set; }
            public bool HideGroupNameRow { get; set; }
            public bool IsEditMode { get; set; }
            public IEnumerable<int> UsedColorIndexes { get; set; }
        }

        public NewRuleDialog(NewRuleDialogOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            serverName = options.Server;
            databaseName = options.Database;
            existingAlias = options.ExistingAlias;
            suggestedName = options.SuggestedName;
            suggestedGroupNameStyle = options.SuggestedGroupNameStyle;
            hideAliasStep = options.HideAliasStep;
            hideGroupNameRow = options.HideGroupNameRow;
            isEditMode = options.IsEditMode;
            usedColorIndexes = new HashSet<int>(options.UsedColorIndexes ?? Enumerable.Empty<int>());

            RuleName = options.SuggestedName;
            SelectedColorIndex = options.SuggestedColorIndex;
            ServerAlias = options.ExistingAlias ?? options.Server;


            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            SizeToContent = SizeToContent.Height;
            Width = 460;
            MinWidth = 460;

            BuildUi(options);
            ApplyThemeResources();
            SetDialogOwner();

            if (!hideAliasStep && string.IsNullOrWhiteSpace(options.ExistingAlias))
            {
                ShowAliasStep();
                txtAlias.Text = options.Server;
            }
            else
            {
                ShowRuleStep();
            }
            EnsureKeyboardCues();
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

        private void SetDialogOwner()
        {
            var appMainWindow = Application.Current?.MainWindow;
            if (appMainWindow != null && !ReferenceEquals(appMainWindow, this))
            {
                Owner = appMainWindow;
                WindowStartupLocation = WindowStartupLocation.CenterOwner;
                return;
            }

            var mainHandle = Process.GetCurrentProcess().MainWindowHandle;
            if (mainHandle != IntPtr.Zero)
            {
                new WindowInteropHelper(this).Owner = mainHandle;
                WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
        }

        private void BuildUi(NewRuleDialogOptions options)
        {
            var root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var contentHost = new Grid
            {
                Margin = new Thickness(16, 14, 16, 10)
            };

            panelAlias = BuildAliasPanel(options.Server, options.Database);
            panelRule = BuildRulePanel(options);

            contentHost.Children.Add(panelAlias);
            contentHost.Children.Add(panelRule);

            panelAliasButtons = BuildAliasButtonPanel();
            panelRuleButtons = BuildRuleButtonPanel();

            var buttonsHost = new Grid
            {
                Margin = new Thickness(0, 4, 0, 0)
            };
            buttonsHost.Children.Add(panelAliasButtons);
            buttonsHost.Children.Add(panelRuleButtons);

            Grid.SetRow(contentHost, 0);
            Grid.SetRow(buttonsHost, 1);

            root.Children.Add(contentHost);
            root.Children.Add(buttonsHost);

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
            ApplyControlBrushes(txtName);
            ApplyControlBrushes(txtAlias);
            ApplyControlBrushes(cmbColor);
            ApplyControlBrushes(btnSave);
            ApplyControlBrushes(btnCancel);
            ApplyControlBrushes(btnOpenConfig);
            ApplyControlBrushes(btnNext);
            ApplyControlBrushes(btnCancelAlias);
            ApplyComboBoxPopupTheme(cmbColor);
            ApplyComboBoxTemplate(cmbColor);
            ApplyDialogButtonStyles();
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

        private void ApplyComboBoxPopupTheme(ComboBox combo)
        {
            if (combo == null)
            {
                return;
            }

            var background = TryFindResource(EnvironmentColors.ToolWindowBackgroundBrushKey) as Brush;
            var foreground = TryFindResource(EnvironmentColors.ToolWindowTextBrushKey) as Brush;
            var highlight = TryFindResource(SystemColors.HighlightBrushKey) as Brush;
            var highlightText = TryFindResource(SystemColors.HighlightTextBrushKey) as Brush;
            var inactiveHighlight = TryFindResource(SystemColors.InactiveSelectionHighlightBrushKey) as Brush;
            var inactiveHighlightText = TryFindResource(SystemColors.InactiveSelectionHighlightTextBrushKey) as Brush;

            if (background != null)
            {
                combo.Resources[SystemColors.WindowBrushKey] = background;
                combo.Resources[SystemColors.ControlBrushKey] = background;
                combo.Resources[SystemColors.InactiveSelectionHighlightBrushKey] = background;
            }

            if (foreground != null)
            {
                combo.Resources[SystemColors.WindowTextBrushKey] = foreground;
                combo.Resources[SystemColors.ControlTextBrushKey] = foreground;
                combo.Resources[SystemColors.InactiveSelectionHighlightTextBrushKey] = foreground;
            }

            if (highlight != null)
            {
                combo.Resources[SystemColors.HighlightBrushKey] = highlight;
            }

            if (highlightText != null)
            {
                combo.Resources[SystemColors.HighlightTextBrushKey] = highlightText;
            }

            if (inactiveHighlight != null)
            {
                combo.Resources[SystemColors.InactiveSelectionHighlightBrushKey] = inactiveHighlight;
            }

            if (inactiveHighlightText != null)
            {
                combo.Resources[SystemColors.InactiveSelectionHighlightTextBrushKey] = inactiveHighlightText;
            }

            var bgColor = GetBrushColor(background, Colors.White);
            bool isLightTheme = GetRelativeLuminance(bgColor) > 0.6;

            if (isLightTheme)
            {
                var hoverColor = BlendColors(bgColor, Colors.Black, 0.06);
                var selectedColor = BlendColors(bgColor, Colors.Black, 0.12);
                combo.Resources["EnvTabsComboHoverBrush"] = new SolidColorBrush(hoverColor);
                combo.Resources["EnvTabsComboSelectedBrush"] = new SolidColorBrush(selectedColor);
                combo.Resources["EnvTabsComboHoverTextBrush"] = new SolidColorBrush(Colors.Black);
                combo.Resources["EnvTabsComboSelectedTextBrush"] = new SolidColorBrush(Colors.Black);
                combo.Resources["EnvTabsComboButtonHoverBrush"] = new SolidColorBrush(BlendColors(bgColor, Colors.Black, 0.08));
                combo.Resources["EnvTabsComboButtonPressedBrush"] = new SolidColorBrush(BlendColors(bgColor, Colors.Black, 0.12));
            }
            else
            {
                var selectedColor = ApplyAlpha(GetBrushColor(highlight, BlendColors(bgColor, Colors.White, 0.18)), 0.78);
                var selectedTextColor = GetBrushColor(highlightText, Colors.White);
                var hoverColor = BlendColors(bgColor, Colors.White, 0.12);
                var hoverTextColor = GetBrushColor(foreground, Colors.White);

                combo.Resources["EnvTabsComboHoverBrush"] = new SolidColorBrush(hoverColor);
                combo.Resources["EnvTabsComboSelectedBrush"] = new SolidColorBrush(selectedColor);
                combo.Resources["EnvTabsComboHoverTextBrush"] = new SolidColorBrush(hoverTextColor);
                combo.Resources["EnvTabsComboSelectedTextBrush"] = new SolidColorBrush(selectedTextColor);
                combo.Resources["EnvTabsComboButtonHoverBrush"] = new SolidColorBrush(BlendColors(bgColor, Colors.White, 0.06));
                combo.Resources["EnvTabsComboButtonPressedBrush"] = new SolidColorBrush(BlendColors(bgColor, Colors.White, 0.10));
            }

            combo.SetResourceReference(Control.BorderBrushProperty, EnvironmentColors.ToolWindowBorderBrushKey);
            combo.BorderThickness = new Thickness(1);
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

        private static Color BlendColors(Color baseColor, Color blendColor, double blendAmount)
        {
            blendAmount = Math.Max(0, Math.Min(1, blendAmount));
            byte r = (byte)(baseColor.R + (blendColor.R - baseColor.R) * blendAmount);
            byte g = (byte)(baseColor.G + (blendColor.G - baseColor.G) * blendAmount);
            byte b = (byte)(baseColor.B + (blendColor.B - baseColor.B) * blendAmount);
            return Color.FromArgb(baseColor.A, r, g, b);
        }

        private static Color ApplyAlpha(Color color, double alphaFactor)
        {
            alphaFactor = Math.Max(0, Math.Min(1, alphaFactor));
            byte a = (byte)(color.A * alphaFactor);
            return Color.FromArgb(a, color.R, color.G, color.B);
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
            var primaryPressed = BlendBrush(accentBrush, Colors.Black, 0.12);

            var accentForeground = isLightTheme ? Brushes.White : Brushes.Black;
            var secondaryForeground = isLightTheme ? Brushes.Black : Brushes.White;

            var baseBorder = WithOpacity(secondaryForeground, 0.28);
            var secondaryHover = isLightTheme
                ? BlendBrush(baseBackground, Colors.Black, 0.06)
                : BlendBrush(baseBackground, Colors.White, 0.06);
            var secondaryPressed = isLightTheme
                ? BlendBrush(baseBackground, Colors.Black, 0.12)
                : BlendBrush(baseBackground, Colors.Black, 0.08);

            var primaryStyle = CreateButtonStyle(accentBrush, baseBorder, accentForeground, primaryHover, primaryPressed);
            var secondaryStyle = CreateButtonStyle(baseBackground, baseBorder, secondaryForeground, secondaryHover, secondaryPressed);

            ApplyButtonStyle(btnSave, primaryStyle);
            ApplyButtonStyle(btnNext, primaryStyle);
            ApplyButtonStyle(btnCancel, secondaryStyle);
            ApplyButtonStyle(btnOpenConfig, secondaryStyle);
            ApplyButtonStyle(btnCancelAlias, secondaryStyle);
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

        private static Brush BlendBrush(Brush brush, Color blendColor, double amount)
        {
            if (brush is SolidColorBrush solid)
            {
                var blended = BlendColors(solid.Color, blendColor, amount);
                return new SolidColorBrush(blended);
            }

            return brush;
        }

        private void ApplyComboBoxTemplate(ComboBox combo)
        {
            if (combo == null)
            {
                return;
            }

            combo.Template = CreateThemedComboBoxTemplate();
        }

        private static ControlTemplate CreateThemedComboBoxTemplate()
        {
            var template = new ControlTemplate(typeof(ComboBox));

            var root = new FrameworkElementFactory(typeof(Grid));
            template.VisualTree = root;

            var toggle = new FrameworkElementFactory(typeof(ToggleButton));
            toggle.Name = "ToggleButton";
            toggle.SetValue(ToggleButton.FocusableProperty, false);
            toggle.SetValue(ToggleButton.ClickModeProperty, ClickMode.Press);
            toggle.SetBinding(Control.BackgroundProperty, new Binding("Background")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });
            toggle.SetBinding(Control.BorderBrushProperty, new Binding("BorderBrush")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });
            toggle.SetBinding(Control.BorderThicknessProperty, new Binding("BorderThickness")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });
            toggle.SetBinding(Control.ForegroundProperty, new Binding("Foreground")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });
            toggle.SetBinding(ToggleButton.IsCheckedProperty, new Binding("IsDropDownOpen")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });

            var toggleTemplate = new ControlTemplate(typeof(ToggleButton));

            var border = new FrameworkElementFactory(typeof(Border));
            border.Name = "Border";
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(2));
            border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Control.BackgroundProperty));
            border.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Control.BorderBrushProperty));
            border.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Control.BorderThicknessProperty));

            var contentDock = new FrameworkElementFactory(typeof(DockPanel));
            contentDock.SetValue(DockPanel.LastChildFillProperty, true);

            var arrow = new FrameworkElementFactory(typeof(Path));
            arrow.Name = "Arrow";
            arrow.SetValue(FrameworkElement.WidthProperty, 8.0);
            arrow.SetValue(FrameworkElement.HeightProperty, 5.0);
            arrow.SetValue(FrameworkElement.MarginProperty, new Thickness(0, 0, 8, 0));
            arrow.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            arrow.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Right);
            arrow.SetValue(DockPanel.DockProperty, Dock.Right);
            arrow.SetValue(Path.DataProperty, Geometry.Parse("M 0 0 L 4 4 L 8 0 Z"));
            arrow.SetValue(Shape.FillProperty, new DynamicResourceExtension(SystemColors.WindowTextBrushKey));

            var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
            contentPresenter.Name = "ContentSite";
            contentPresenter.SetValue(ContentPresenter.MarginProperty, new Thickness(6, 2, 4, 2));
            contentPresenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Left);
            contentPresenter.SetValue(ContentPresenter.RecognizesAccessKeyProperty, true);
            contentPresenter.SetBinding(ContentPresenter.ContentProperty, new Binding("SelectedItem")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(ComboBox), 1)
            });
            contentPresenter.SetBinding(ContentPresenter.ContentTemplateProperty, new Binding("ItemTemplate")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(ComboBox), 1)
            });
            contentPresenter.SetBinding(ContentPresenter.ContentTemplateSelectorProperty, new Binding("ItemTemplateSelector")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(ComboBox), 1)
            });
            contentPresenter.SetBinding(TextElement.ForegroundProperty, new Binding("Foreground")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(ComboBox), 1)
            });

            contentDock.AppendChild(arrow);
            contentDock.AppendChild(contentPresenter);
            border.AppendChild(contentDock);
            toggleTemplate.VisualTree = border;

            var toggleHoverTrigger = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            toggleHoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, new DynamicResourceExtension("EnvTabsComboButtonHoverBrush"), "Border"));
            toggleTemplate.Triggers.Add(toggleHoverTrigger);

            var toggleCheckedTrigger = new Trigger { Property = ToggleButton.IsCheckedProperty, Value = true };
            toggleCheckedTrigger.Setters.Add(new Setter(Border.BackgroundProperty, new DynamicResourceExtension("EnvTabsComboButtonPressedBrush"), "Border"));
            toggleTemplate.Triggers.Add(toggleCheckedTrigger);

            toggle.SetValue(Control.TemplateProperty, toggleTemplate);
            root.AppendChild(toggle);

            var popup = new FrameworkElementFactory(typeof(Popup));
            popup.Name = "PART_Popup";
            popup.SetValue(Popup.PlacementProperty, PlacementMode.Bottom);
            popup.SetValue(Popup.AllowsTransparencyProperty, true);
            popup.SetValue(Popup.FocusableProperty, false);
            popup.SetBinding(Popup.IsOpenProperty, new Binding("IsDropDownOpen")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });

            var popupBorder = new FrameworkElementFactory(typeof(Border));
            popupBorder.SetValue(Border.BackgroundProperty, new DynamicResourceExtension(SystemColors.WindowBrushKey));
            popupBorder.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Control.BorderBrushProperty));
            popupBorder.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Control.BorderThicknessProperty));
            popupBorder.SetValue(Border.SnapsToDevicePixelsProperty, true);

            var scrollViewer = new FrameworkElementFactory(typeof(ScrollViewer));
            scrollViewer.SetValue(ScrollViewer.SnapsToDevicePixelsProperty, true);
            scrollViewer.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
            scrollViewer.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
            scrollViewer.SetBinding(ScrollViewer.CanContentScrollProperty, new Binding("CanContentScroll")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });

            var itemsPresenter = new FrameworkElementFactory(typeof(ItemsPresenter));
            scrollViewer.AppendChild(itemsPresenter);
            popupBorder.AppendChild(scrollViewer);
            popup.AppendChild(popupBorder);
            root.AppendChild(popup);

            var disabledTrigger = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
            disabledTrigger.Setters.Add(new Setter(Control.ForegroundProperty, new DynamicResourceExtension(SystemColors.GrayTextBrushKey)));
            template.Triggers.Add(disabledTrigger);

            return template;
        }

        private Grid BuildAliasPanel(string server, string database)
        {
            var panel = new Grid
            {
                Visibility = Visibility.Collapsed
            };

            var stack = new StackPanel { Orientation = Orientation.Vertical };

            lblAliasHeader = new TextBlock
            {
                Text = "Assign an alias for this server",
                FontWeight = FontWeights.Bold,
                FontSize = SystemFonts.MessageFontSize + 5,
                Margin = new Thickness(0, 0, 0, 12)
            };

            stack.Children.Add(lblAliasHeader);
            stack.Children.Add(BuildConnectionLine(server, database, includeDatabase: false));

            var grid = CreateFormGrid();
            int row = 0;

            txtAlias = new TextBox
            {
                Width = 270,
                MinHeight = 26,
                Margin = new Thickness(0, 2, 0, 2),
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(6, 2, 6, 2)
            };

            AddLabelValueRow(grid, row++, "Alias", txtAlias);

            stack.Children.Add(grid);
            panel.Children.Add(stack);

            return panel;
        }

        private Grid BuildRulePanel(NewRuleDialogOptions options)
        {
            var panel = new Grid
            {
                Visibility = Visibility.Collapsed
            };

            var stack = new StackPanel { Orientation = Orientation.Vertical };

            lblHeader = new TextBlock
            {
                Text = "Name and color tabs for this connection",
                FontWeight = FontWeights.Bold,
                FontSize = SystemFonts.MessageFontSize + 5,
                Margin = new Thickness(0, 0, 0, 12),
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 420
            };

            stack.Children.Add(lblHeader);
            stack.Children.Add(BuildConnectionLine(options.Server, options.HideDatabaseRow ? null : options.Database));

            var ruleGrid = CreateFormGrid();
            int row = 0;

            if (!hideGroupNameRow)
            {
                txtName = new TextBox
                {
                    Width = 270,
                    MinHeight = 26,
                    Margin = new Thickness(0, 2, 0, 2),
                    Text = options.SuggestedName,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    Padding = new Thickness(6, 2, 6, 2)
                };
                AddLabelValueRow(ruleGrid, row++, "Group Name", txtName);
            }

            cmbColor = new ComboBox
            {
                Width = 270,
                MinHeight = 26,
                IsEditable = false,
                Margin = new Thickness(0, 2, 0, 2),
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(6, 2, 6, 2)
            };

            cmbColor.ItemTemplate = CreateColorItemTemplate();
            cmbColor.ItemContainerStyle = CreateColorItemContainerStyle();

            var orderedList = BuildOrderedColorList(options.SuggestedColorIndex);
            cmbColor.ItemsSource = orderedList;
            cmbColor.SelectedItem = orderedList.FirstOrDefault(c => c.Index == options.SuggestedColorIndex) ?? orderedList.FirstOrDefault();

            AddLabelValueRow(ruleGrid, row++, "Color", cmbColor);

            stack.Children.Add(ruleGrid);
            panel.Children.Add(stack);

            return panel;
        }

        private StackPanel BuildRuleButtonPanel()
        {
            btnSave = new Button
            {
                Content = "_Save",
                MinWidth = 110,
                Height = 26,
                IsDefault = true,
                Margin = new Thickness(6, 10, 6, 10)
            };
            btnSave.Click += (s, e) =>
            {
                RuleName = txtName != null ? (string.IsNullOrWhiteSpace(txtName.Text) ? null : txtName.Text.Trim()) : suggestedName;
                if (cmbColor.SelectedItem is ColorItem item)
                {
                    SelectedColorIndex = item.Index; // null for the "None" item
                }
                dialogResult = WinFormsDialogResult.OK;
                this.DialogResult = true;
            };

            btnCancel = new Button
            {
                Content = "_Cancel",
                MinWidth = 110,
                Height = 26,
                IsCancel = true,
                Margin = new Thickness(6, 10, 6, 10)
            };
            btnCancel.Click += (s, e) =>
            {
                // Cancel means "I know about this connection but don't want to rename or color it".
                // We set null values so AutoConfigurationService can save a silent rule.
                RuleName = null;
                SelectedColorIndex = null;
                dialogResult = WinFormsDialogResult.Cancel;
                this.DialogResult = false;
            };

            btnOpenConfig = new Button
            {
                Content = "_Open Config",
                MinWidth = 120,
                Height = 26,
                Margin = new Thickness(6, 10, 6, 10)
            };
            btnOpenConfig.Click += (s, e) =>
            {
                RuleName = txtName != null ? (string.IsNullOrWhiteSpace(txtName.Text) ? null : txtName.Text.Trim()) : suggestedName;
                if (cmbColor.SelectedItem is ColorItem item)
                {
                    SelectedColorIndex = item.Index; // null for the "None" item
                }
                OpenConfigRequested = true;
                dialogResult = WinFormsDialogResult.Yes;
                this.DialogResult = true;
            };

            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            panel.Children.Add(btnSave);
            panel.Children.Add(btnCancel);
            panel.Children.Add(btnOpenConfig);

            return panel;
        }

        private StackPanel BuildAliasButtonPanel()
        {
            btnNext = new Button
            {
                Content = "_Next >",
                MinWidth = 110,
                Height = 26,
                IsDefault = true,
                Margin = new Thickness(6, 10, 6, 10)
            };
            btnNext.Click += (s, e) =>
            {
                ServerAlias = string.IsNullOrWhiteSpace(txtAlias.Text) ? serverName : txtAlias.Text.Trim();
                AliasConfirmed?.Invoke(ServerAlias);
                ShowRuleStep();
            };

            btnCancelAlias = new Button
            {
                Content = "_Cancel",
                MinWidth = 110,
                Height = 26,
                Margin = new Thickness(6, 10, 6, 10)
            };
            btnCancelAlias.Click += (s, e) =>
            {
                ServerAlias = serverName;
                ShowRuleStep();
            };

            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            panel.Children.Add(btnNext);
            panel.Children.Add(btnCancelAlias);

            return panel;
        }

        private void ShowAliasStep()
        {
            Title = "SSMS EnvTabs - Assign Server Alias";
            panelAlias.Visibility = Visibility.Visible;
            panelRule.Visibility = Visibility.Collapsed;
            panelAliasButtons.Visibility = Visibility.Visible;
            panelRuleButtons.Visibility = Visibility.Collapsed;
            btnNext.IsDefault = true;
            btnCancelAlias.IsCancel = true;
            txtAlias?.Focus();
            EnsureKeyboardCues();
        }

        private void ShowRuleStep()
        {
            Title = isEditMode ? "SSMS EnvTabs - Edit Rule" : "SSMS EnvTabs - New Rule";
            panelAlias.Visibility = Visibility.Collapsed;
            panelRule.Visibility = Visibility.Visible;
            panelAliasButtons.Visibility = Visibility.Collapsed;
            panelRuleButtons.Visibility = Visibility.Visible;

            if (!isEditMode && !hideGroupNameRow && txtName != null)
            {
                string expectedWithServer = BuildSuggestedGroupName(serverName, serverName, databaseName);
                string expectedWithExistingAlias = BuildSuggestedGroupName(serverName, string.IsNullOrWhiteSpace(existingAlias) ? serverName : existingAlias, databaseName);
                string expectedWithAlias = BuildSuggestedGroupName(serverName, string.IsNullOrWhiteSpace(ServerAlias) ? serverName : ServerAlias, databaseName);

                bool matchesSuggested = string.IsNullOrWhiteSpace(txtName.Text)
                    || string.Equals(txtName.Text, expectedWithServer, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(txtName.Text, expectedWithExistingAlias, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(txtName.Text, suggestedName, StringComparison.OrdinalIgnoreCase);

                if (matchesSuggested)
                {
                    txtName.Text = expectedWithAlias;
                }
            }

            btnSave.IsDefault = true;
            btnCancel.IsCancel = true;
            if (txtName != null)
            {
                txtName.Focus();
            }
            else
            {
                cmbColor?.Focus();
            }
            EnsureKeyboardCues();
        }

        private void EnsureKeyboardCues()
        {
            TrySetKeyboardCues(this);

            if (Content is DependencyObject root)
            {
                TrySetKeyboardCues(root);
            }

            panelAlias?.InvalidateVisual();
            panelRule?.InvalidateVisual();
            panelAliasButtons?.InvalidateVisual();
            panelRuleButtons?.InvalidateVisual();
        }

        private static void TrySetKeyboardCues(DependencyObject target)
        {
            var dp = typeof(KeyboardNavigation).GetField("ShowKeyboardCuesProperty", BindingFlags.Public | BindingFlags.Static)
                ?.GetValue(null) as DependencyProperty
                ?? typeof(AccessText).GetField("ShowKeyboardCuesProperty", BindingFlags.Public | BindingFlags.Static)
                ?.GetValue(null) as DependencyProperty;

            if (dp != null)
            {
                target.SetValue(dp, true);
            }
        }

        private string BuildSuggestedGroupName(string serverValue, string serverAliasValue, string databaseValue)
        {
            string serverPart = serverValue ?? string.Empty;
            string aliasPart = !string.IsNullOrEmpty(serverAliasValue) ? serverAliasValue : serverPart;
            string dbPart = databaseValue ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(suggestedGroupNameStyle))
            {
                return suggestedGroupNameStyle
                    .Replace("[server]", serverPart)
                    .Replace("[serverAlias]", aliasPart)
                    .Replace("[db]", dbPart);
            }

            if (string.IsNullOrWhiteSpace(dbPart))
            {
                return serverPart;
            }

            return $"{serverPart} {dbPart}";
        }

        private List<ColorItem> BuildOrderedColorList(int? suggestedColorIndex)
        {
            var orderedList = new List<ColorItem>();
            var suggestedItem = ColorList.FirstOrDefault(c => c.Index == suggestedColorIndex);

            foreach (var item in ColorList)
            {
                bool isUsed = item.Index.HasValue && usedColorIndexes.Contains(item.Index.Value);
                item.DisplayName = isUsed ? $"{item.Name} (used)" : item.Name;
            }

            if (suggestedItem != null)
            {
                orderedList.Add(suggestedItem);
                orderedList.AddRange(ColorList.Where(c => c.Index != suggestedColorIndex));
            }
            else
            {
                orderedList.AddRange(ColorList);
            }

            return orderedList;
        }

        private static Grid CreateFormGrid()
        {
            var grid = new Grid
            {
                Margin = new Thickness(0, 0, 0, 6)
            };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            return grid;
        }

        private static void AddLabelValueRow(Grid grid, int rowIndex, string labelText, UIElement valueElement)
        {
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var label = new TextBlock
            {
                Text = labelText,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 4, 12, 4),
                HorizontalAlignment = HorizontalAlignment.Right,
                TextAlignment = TextAlignment.Right
            };

            Grid.SetRow(label, rowIndex);
            Grid.SetColumn(label, 0);

            if (valueElement is FrameworkElement frameworkElement)
            {
                frameworkElement.VerticalAlignment = VerticalAlignment.Center;
            }

            Grid.SetRow(valueElement, rowIndex);
            Grid.SetColumn(valueElement, 1);

            grid.Children.Add(label);
            grid.Children.Add(valueElement);
        }

        private UIElement BuildConnectionLine(string server, string database, bool includeDatabase = true)
        {
            string displayText = !includeDatabase || string.IsNullOrWhiteSpace(database)
                ? server
                : $"{server}, {database}";

            var line = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 12)
            };

            line.Children.Add(CreateDatabaseIcon(18));
            line.Children.Add(new TextBlock
            {
                Text = displayText,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            });

            return line;
        }

        private UIElement CreateDatabaseIcon(double size)
        {
            var stroke = TryFindResource(EnvironmentColors.ToolWindowTextBrushKey) as Brush
                ?? Brushes.Gray;
            var fill = WithOpacity(stroke, 1);

            const double canvasSize = 16;
            
            // Metrics for a cylinder icon
            const double glyphLeft = 2;
            const double glyphWidth = 12;
            const double glyphRight = glyphLeft + glyphWidth;
            
            const double ellipseHeight = 4;
            const double ellipseHalfHeight = ellipseHeight / 2;
            
            const double topY = 1;
            const double cylinderHeight = 9; // Height from top-center to bottom-center
            
            double sideTopY = topY + ellipseHalfHeight;
            double sideBottomY = sideTopY + cylinderHeight;

            var canvas = new Canvas
            {
                Width = canvasSize,
                Height = canvasSize
            };

            // 1. Body Path: Sides and Bottom Curve
            // Using a Path instead of Rectangle avoids the "flat bottom" look and allows
            // us to stroke only lines and the bottom arc, leaving the top un-stroked.
            var bodyPath = new Path
            {
                Stroke = stroke,
                //Fill = fill,
                StrokeThickness = 1
            };
            
            var geometry = new PathGeometry();
            var figure = new PathFigure
            {
                StartPoint = new Point(glyphLeft, sideTopY),
                IsClosed = false // Open at the top so we don't draw a stroke there
            };
            
            // Left vertical line
            figure.Segments.Add(new LineSegment(new Point(glyphLeft, sideBottomY), true));
            
            // Bottom arc (bulging down)
            figure.Segments.Add(new ArcSegment(
                new Point(glyphRight, sideBottomY),
                new Size(glyphWidth / 2, ellipseHalfHeight),
                0,
                false, 
                SweepDirection.Counterclockwise,
                true));
            
            // Right vertical line
            figure.Segments.Add(new LineSegment(new Point(glyphRight, sideTopY), true));
            
            geometry.Figures.Add(figure);
            bodyPath.Data = geometry;
            canvas.Children.Add(bodyPath);

            // 2. Top Ellipse (Lid)
            // Draws full ellipse on top, creating the "closed cylinder" look
            var top = new Ellipse
            {
                Width = glyphWidth,
                Height = ellipseHeight,
                Stroke = stroke,
                Fill = fill,
                StrokeThickness = 1
            };
            Canvas.SetLeft(top, glyphLeft);
            Canvas.SetTop(top, topY);
            canvas.Children.Add(top);

            return new Viewbox
            {
                Width = size,
                Height = size,
                Stretch = Stretch.Uniform,
                Margin = new Thickness(0, 0, 8, 0),
                Child = canvas
            };
        }

        private static DataTemplate CreateColorItemTemplate()
        {
            var stackFactory = new FrameworkElementFactory(typeof(StackPanel));
            stackFactory.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);

            var swatchFactory = new FrameworkElementFactory(typeof(Border));
            swatchFactory.SetValue(Border.WidthProperty, 18.0);
            swatchFactory.SetValue(Border.HeightProperty, 16.0);
            swatchFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(3));
            swatchFactory.SetValue(Border.MarginProperty, new Thickness(0, 0, 8, 0));
            swatchFactory.SetBinding(Border.BackgroundProperty, new Binding("SwatchBrush"));

            var textFactory = new FrameworkElementFactory(typeof(TextBlock));
            textFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            textFactory.SetBinding(TextBlock.TextProperty, new Binding("DisplayName"));

            stackFactory.AppendChild(swatchFactory);
            stackFactory.AppendChild(textFactory);

            return new DataTemplate { VisualTree = stackFactory };
        }

        private static Style CreateColorItemContainerStyle()
        {
            var style = new Style(typeof(ComboBoxItem));
            style.Setters.Add(new Setter(ComboBoxItem.MinHeightProperty, 22.0));
            style.Setters.Add(new Setter(ComboBoxItem.PaddingProperty, new Thickness(4, 2, 4, 2)));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new DynamicResourceExtension(EnvironmentColors.ToolWindowBackgroundBrushKey)));
            style.Setters.Add(new Setter(Control.ForegroundProperty, new DynamicResourceExtension(EnvironmentColors.ToolWindowTextBrushKey)));

            var template = new ControlTemplate(typeof(ComboBoxItem));
            var border = new FrameworkElementFactory(typeof(Border));
            border.Name = "Bd";
            border.SetValue(Border.SnapsToDevicePixelsProperty, true);
            border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Control.BackgroundProperty));

            var presenter = new FrameworkElementFactory(typeof(ContentPresenter));
            presenter.SetValue(ContentPresenter.MarginProperty, new Thickness(2, 0, 2, 0));
            presenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            presenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Left);
            presenter.SetValue(ContentPresenter.SnapsToDevicePixelsProperty, true);
            presenter.SetValue(TextElement.ForegroundProperty, new TemplateBindingExtension(Control.ForegroundProperty));

            border.AppendChild(presenter);
            template.VisualTree = border;

            var highlightedTrigger = new Trigger { Property = ComboBoxItem.IsHighlightedProperty, Value = true };
            highlightedTrigger.Setters.Add(new Setter(Control.BackgroundProperty, new DynamicResourceExtension("EnvTabsComboHoverBrush"), "Bd"));
            highlightedTrigger.Setters.Add(new Setter(Control.ForegroundProperty, new DynamicResourceExtension("EnvTabsComboHoverTextBrush")));
            template.Triggers.Add(highlightedTrigger);

            var disabledTrigger = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
            disabledTrigger.Setters.Add(new Setter(Control.ForegroundProperty, new DynamicResourceExtension(SystemColors.GrayTextBrushKey)));
            template.Triggers.Add(disabledTrigger);

            style.Setters.Add(new Setter(Control.TemplateProperty, template));
            return style;
        }
    }
}
