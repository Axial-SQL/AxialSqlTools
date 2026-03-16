namespace AxialSqlTools
{
    using System;
    using System.Reflection;
    using System.Windows;
    using System.Windows.Media;
    using Microsoft.VisualStudio.PlatformUI;

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
    }
}