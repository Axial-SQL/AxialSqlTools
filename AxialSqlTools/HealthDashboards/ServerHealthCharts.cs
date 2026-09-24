using ScottPlot;
using ScottPlot.Plottables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

namespace AxialSqlTools
{
    internal sealed class ServerHealthChartTheme
    {
        internal Color Background, Foreground, Border, Grid;
        internal Color[] Series;
        internal bool HighContrast;
        internal Color GetColor(int index) => Series[index % Series.Length];
    }

    // No WPF dependencies: the same charts can be rendered and checked off-screen.
    internal static class ServerHealthCharts
    {
        private sealed class ChartState
        {
            internal readonly List<Action<ServerHealthChartTheme>> ApplyColors = new List<Action<ServerHealthChartTheme>>();
            internal Func<Coordinates, string> Hover;
            internal readonly HashSet<string> HiddenSeries = new HashSet<string>(StringComparer.Ordinal);
            internal Action VisibilityChanged;
            internal ScottPlot.Panels.LegendPanel LegendPanel;
            internal int LegendRender = -1;
            internal readonly List<Tuple<PixelRect, LegendItem>> LegendHits = new List<Tuple<PixelRect, LegendItem>>();
        }

        private static readonly ConditionalWeakTable<Plot, ChartState> States = new ConditionalWeakTable<Plot, ChartState>();

        private sealed class LegendAwareLayout : ILayoutEngine
        {
            private readonly ScottPlot.Panels.LegendPanel _panel;
            private readonly ScottPlot.LayoutEngines.Automatic _automatic = new ScottPlot.LayoutEngines.Automatic();

            internal LegendAwareLayout(ScottPlot.Panels.LegendPanel panel) => _panel = panel;

            public Layout GetLayout(PixelRect figureRect, Plot plot, Paint paint)
            {
                // The default panel measures the previous legend height. Measure all wrapped
                // rows at this render's available width so entries cannot be clipped on resize.
                _panel.MinimumSize = _panel.MaximumSize = 0;
                var initial = _automatic.GetLayout(figureRect, plot, paint);
                if (plot.Legend.GetItems().Length > 0)
                {
                    var legend = plot.Legend.GetLayout(new PixelSize(Math.Max(1, initial.DataRect.Width), float.PositiveInfinity), paint);
                    _panel.MinimumSize = _panel.MaximumSize = legend.LegendRect.Height + _panel.Padding.Vertical;
                }
                return _automatic.GetLayout(figureRect, plot, paint);
            }
        }

        internal static Plot Create(string title, string units = null)
        {
            var plot = new Plot();
            plot.Title(title, 14);
            plot.YLabel(units ?? "", 11);
            plot.Axes.Left.TickLabelStyle.FontSize = 11;
            plot.Axes.Bottom.TickLabelStyle.FontSize = 11;
            // Leave room for the last time label at the edge of a small dashboard tile.
            plot.Axes.Right.MinimumSize = 28;
            plot.Axes.Right.TickGenerator = new ScottPlot.TickGenerators.NumericManual();
            plot.Axes.Right.FrameLineStyle.Width = 0;
            plot.Axes.Top.IsVisible = false;
            plot.HideLegend();
            plot.Legend.FontSize = 11;
            plot.Legend.Padding = new PixelPadding(3);
            plot.Legend.InterItemPadding = new PixelPadding(4, 2);
            plot.Legend.ShadowColor = Colors.Transparent;
            plot.Legend.TightHorizontalWrapping = true;
            plot.Legend.ShowItemsFromHiddenPlottables = true;
            plot.Legend.HiddenItemOpacity = 0.45;
            return plot;
        }

        internal static void ApplyTheme(Plot plot, ServerHealthChartTheme theme)
        {
            plot.FigureBackground.Color = theme.Background;
            plot.DataBackground.Color = theme.Background;
            plot.Axes.Color(theme.Foreground);
            plot.Axes.Title.Label.ForeColor = theme.Foreground;
            plot.Grid.MajorLineColor = theme.Grid;
            plot.Grid.MinorLineColor = theme.Grid;
            plot.Legend.FontColor = theme.Foreground;
            plot.Legend.BackgroundColor = theme.Background;
            plot.Legend.OutlineColor = theme.Border;
            foreach (var apply in States.GetOrCreateValue(plot).ApplyColors) apply(theme);
        }

        internal static void SetHover(Plot plot, Func<Coordinates, string> hover) => States.GetOrCreateValue(plot).Hover = hover;
        internal static string HoverText(Plot plot, Coordinates point) => States.GetOrCreateValue(plot).Hover?.Invoke(point);

        internal static bool IsSeriesVisible(Plot plot, string label) => !States.GetOrCreateValue(plot).HiddenSeries.Contains(label);

        internal static void CopyVisibility(Plot previous, Plot replacement)
        {
            var state = States.GetOrCreateValue(replacement);
            state.HiddenSeries.UnionWith(States.GetOrCreateValue(previous).HiddenSeries);
            foreach (var item in replacement.Legend.GetItems())
                if (item.Plottable != null) item.Plottable.IsVisible = !state.HiddenSeries.Contains(item.LabelText);
            state.VisibilityChanged?.Invoke();
        }

        internal static bool IsOverData(Plot plot, Pixel pixel) => plot.RenderManager.LastRender.Count > 0
            && plot.RenderManager.LastRender.DataRect.Contains(pixel.Divide((float)plot.ScaleFactor));

        internal static LegendItem LegendItemAt(Plot plot, Pixel pixel)
        {
            var state = States.GetOrCreateValue(plot);
            var render = plot.RenderManager.LastRender;
            var panel = state.LegendPanel;
            if (panel == null || !panel.IsVisible || render.Count == 0) return null;
            // Use the rendered panel's size and alignment, including wrapped labels and DPI scaling.
            if (state.LegendRender != render.Count)
            {
                state.LegendHits.Clear();
                state.LegendRender = render.Count;
                if (!render.Layout.PanelSizes.TryGetValue(panel, out float size)
                    || !render.Layout.PanelOffsets.TryGetValue(panel, out float offset)) return null;
                using (var paint = Paint.NewDisposablePaint())
                {
                    var rect = panel.GetPanelRect(render.DataRect, size, offset, paint);
                    var layout = plot.Legend.GetLayout(rect.Size, paint);
                    var aligned = layout.LegendRect.AlignedInside(rect, panel.Alignment);
                    var shift = new PixelOffset(aligned.Left, aligned.Top);
                    for (int i = 0; i < layout.LegendItems.Length; i++)
                    {
                        var item = layout.LegendItems[i];
                        if (item.Plottable == null) continue;
                        var hit = new PixelRect(layout.SymbolRects[i].TopLeft, layout.LabelRects[i].BottomRight).WithOffset(shift);
                        state.LegendHits.Add(Tuple.Create(hit, item));
                    }
                }
            }
            var point = pixel.Divide((float)plot.ScaleFactor);
            return state.LegendHits.FirstOrDefault(x => x.Item1.Contains(point))?.Item2;
        }

        internal static bool ToggleLegendAt(Plot plot, Pixel pixel)
        {
            var item = LegendItemAt(plot, pixel);
            if (item == null) return false;
            var state = States.GetOrCreateValue(plot);
            item.Plottable.IsVisible = !item.Plottable.IsVisible;
            if (item.Plottable.IsVisible) state.HiddenSeries.Remove(item.LabelText);
            else state.HiddenSeries.Add(item.LabelText);
            state.VisibilityChanged?.Invoke();
            return true;
        }

        internal static void EnableStackToggling(Plot plot, bool horizontal = false, double headroom = 1.1)
        {
            var series = plot.GetPlottables<BarPlot>().ToArray();
            var bars = series.Select(x => x.Bars.ToArray()).ToArray();
            var values = bars.Select(group => group.Select(bar => bar.Value - bar.ValueBase).ToArray()).ToArray();
            States.GetOrCreateValue(plot).VisibilityChanged = () =>
            {
                var totals = new Dictionary<double, double>();
                for (int s = 0; s < series.Length; s++)
                    for (int b = 0; b < bars[s].Length; b++)
                    {
                        var bar = bars[s][b];
                        totals.TryGetValue(bar.Position, out double baseline);
                        bar.ValueBase = baseline;
                        bar.Value = baseline + values[s][b];
                        if (series[s].IsVisible) totals[bar.Position] = bar.Value;
                    }
                double maximum = Math.Max(1, totals.Values.DefaultIfEmpty(0).Max() * headroom);
                if (horizontal) plot.Axes.SetLimitsX(0, maximum);
                else plot.Axes.SetLimitsY(0, maximum);
            };
        }

        internal static void LegendBelow(Plot plot)
        {
            var panel = plot.ShowLegend(Edge.Bottom);
            panel.Padding = new PixelPadding(0, 2);
            States.GetOrCreateValue(plot).LegendPanel = panel;
            plot.Layout.LayoutEngine = new LegendAwareLayout(panel);
        }

        internal static void TimeAxis(Plot plot, string format = "HH:mm")
        {
            plot.Axes.DateTimeTicksBottom();
            plot.Axes.Bottom.TickGenerator = new ScottPlot.TickGenerators.DateTimeAutomatic
            {
                LabelFormatter = time => time.ToString(format)
            };
            plot.Axes.Bottom.TickLabelStyle.FontSize = 11;
        }

        internal static Scatter Line(Plot plot, double[] xs, double[] ys, string label, int colorIndex)
        {
            var line = plot.Add.Scatter(xs, ys);
            line.LegendText = label;
            line.LineWidth = 1.8f;
            line.MarkerSize = xs.Length == 1 ? 4 : 0;
            line.LinePattern = LinePattern.Solid;
            States.GetOrCreateValue(plot).ApplyColors.Add(theme => line.Color = theme.GetColor(colorIndex));
            return line;
        }

        internal static BarPlot Bars(Plot plot, List<Bar> bars, string label, int colorIndex)
        {
            var series = plot.Add.Bars(bars);
            // Manual legend uses the same fill style so high-contrast hatching remains identifiable.
            if (!string.IsNullOrEmpty(label) && bars.Count > 0)
                plot.Legend.ManualItems.Add(new LegendItem { LabelText = label, FillStyle = bars[0].FillStyle, Plottable = series });
            States.GetOrCreateValue(plot).ApplyColors.Add(theme =>
            {
                series.ValueLabelStyle.ForeColor = theme.HighContrast ? theme.Foreground : theme.Background;
                foreach (Bar bar in bars)
                {
                    bar.FillColor = theme.HighContrast ? theme.Background : theme.GetColor(colorIndex);
                    bar.LineColor = theme.HighContrast ? theme.Foreground : theme.Border;
                    bar.FillHatchColor = theme.Foreground;
                    bar.FillHatch = theme.HighContrast ? Hatch(colorIndex) : null;
                }
            });
            return series;
        }

        private static IHatch Hatch(int index)
        {
            switch (index % 7)
            {
                case 0: return new ScottPlot.Hatches.Striped(ScottPlot.Hatches.StripeDirection.DiagonalUp);
                case 1: return new ScottPlot.Hatches.Striped(ScottPlot.Hatches.StripeDirection.DiagonalDown);
                case 2: return new ScottPlot.Hatches.Striped(ScottPlot.Hatches.StripeDirection.Horizontal);
                case 3: return new ScottPlot.Hatches.Striped(ScottPlot.Hatches.StripeDirection.Vertical);
                case 4: return new ScottPlot.Hatches.Dots();
                case 5: return new ScottPlot.Hatches.Grid();
                default: return new ScottPlot.Hatches.Checker();
            }
        }

        internal static Plot Waits(KeyValuePair<DateTime, Dictionary<string, decimal>>[] minutes, DateTime now, Dictionary<string, int> colors)
        {
            var plot = Create("Real-time waits · 1-minute buckets", "Wait seconds");
            TimeAxis(plot);
            // Rank by recent deltas, not lifetime totals. Retain every remaining wait in Other.
            var ranked = minutes.SelectMany(x => x.Value).GroupBy(x => x.Key)
                .Select(g => new { Name = g.Key, Total = g.Sum(x => x.Value) })
                .Where(x => x.Total > 0).OrderByDescending(x => x.Total).ThenBy(x => x.Name, StringComparer.Ordinal)
                .ToArray();
            string[] top = ranked.Take(6).Select(x => x.Name).ToArray();
            foreach (string old in colors.Keys.Except(top).ToArray()) colors.Remove(old);
            foreach (string name in top.Where(name => !colors.ContainsKey(name)))
                colors[name] = Enumerable.Range(0, 6).First(index => !colors.Values.Contains(index));
            var groups = ranked.Take(6).Select(x => new
            {
                Label = x.Name, Names = new[] { x.Name }, x.Total, ColorIndex = colors[x.Name]
            }).ToList();
            if (ranked.Length > top.Length) groups.Add(new
            {
                Label = "Other waits", Names = ranked.Skip(6).Select(x => x.Name).ToArray(),
                Total = ranked.Skip(6).Sum(x => x.Total), ColorIndex = 6
            });
            // Keep stack and legend order aligned by total seconds, including the combined Other group.
            groups = groups.OrderByDescending(x => x.Total).ThenBy(x => x.Label, StringComparer.Ordinal).ToList();
            var bases = new double[minutes.Length];
            foreach (var group in groups)
            {
                var bars = new List<Bar>();
                for (int i = 0; i < minutes.Length; i++)
                {
                    double value = (double)group.Names.Sum(name => minutes[i].Value.TryGetValue(name, out decimal wait) ? wait : 0);
                    bars.Add(new Bar { Position = minutes[i].Key.AddSeconds(30).ToOADate(), Size = 50.0 / 86400,
                        ValueBase = bases[i], Value = bases[i] + value, LineWidth = 0.5f });
                    bases[i] += value;
                }
                // Keep each visible wait's color when rankings change; assign new waits unused colors.
                Bars(plot, bars, group.Label, group.ColorIndex);
            }
            if (groups.Count > 0) LegendBelow(plot);
            else EmptyMessage(plot, minutes.Length == 0 ? "Collecting wait samples..." : "No non-idle waits recorded");
            DateTime currentMinute = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0);
            plot.Axes.SetLimitsX(currentMinute.AddMinutes(-14).ToOADate(), currentMinute.AddMinutes(1).ToOADate());
            plot.Axes.SetLimitsY(0, Math.Max(1, bases.DefaultIfEmpty(0).Max() * 1.1));
            EnableStackToggling(plot);
            SetHover(plot, point =>
            {
                var bucket = minutes.FirstOrDefault(x => point.X >= x.Key.ToOADate() && point.X < x.Key.AddMinutes(1).ToOADate());
                if (bucket.Value == null) return "Collecting wait samples...";
                var visible = groups.Where(x => IsSeriesVisible(plot, x.Label)).Select(group => new
                {
                    group.Label, Seconds = group.Names.Sum(name => bucket.Value.TryGetValue(name, out decimal value) ? value : 0)
                }).OrderByDescending(x => x.Seconds).ThenBy(x => x.Label, StringComparer.Ordinal).ToArray();
                if (visible.Length == 0) return null;
                return bucket.Key.ToString("HH:mm") + " · wait seconds\n" + string.Join("\n", visible.Select(x =>
                    x.Label + ": " + x.Seconds.ToString("N2")));
            });
            return plot;
        }

        internal static Plot TimeSeries(string title, string units, string[] labels, double[] xs,
            double[][] values, DateTime now, bool percent = false)
        {
            var plot = Create(title, units);
            TimeAxis(plot);
            for (int i = 0; i < values.Length; i++)
                if (xs.Length > 0) Line(plot, xs, values[i], labels[i], i);
            if (values.Length > 1 && xs.Length > 0) LegendBelow(plot);
            plot.Axes.SetLimitsX(now.AddMinutes(-15).ToOADate(), now.ToOADate());
            Action scaleVisible = () =>
            {
                double maximum = values.Where((ys, i) => IsSeriesVisible(plot, labels[i])).SelectMany(x => x)
                    .Where(x => !double.IsNaN(x) && !double.IsInfinity(x)).DefaultIfEmpty(0).Max();
                plot.Axes.SetLimitsY(0, percent ? 100 : Math.Max(1, maximum * 1.1));
            };
            States.GetOrCreateValue(plot).VisibilityChanged = scaleVisible;
            scaleVisible();
            if (xs.Length == 0) EmptyMessage(plot, "Collecting samples...");
            SetHover(plot, point =>
            {
                if (xs.Length == 0) return "Collecting samples...";
                // Clamp to the nearest recorded sample, including a chart's first single sample.
                int index = Array.BinarySearch(xs, point.X);
                if (index < 0) index = Math.Min(~index, xs.Length - 1);
                if (index > 0 && Math.Abs(xs[index - 1] - point.X) < Math.Abs(xs[index] - point.X)) index--;
                var visible = Enumerable.Range(0, values.Length).Where(i => IsSeriesVisible(plot, labels[i])).ToArray();
                if (visible.Length == 0) return null;
                return DateTime.FromOADate(xs[index]).ToString("HH:mm:ss") + "\n" + string.Join("\n", visible.Select(i =>
                    labels[i] + ": " + (double.IsNaN(values[i][index]) ? "Collecting..." : values[i][index].ToString("N1") + " " + units)));
            });
            return plot;
        }

        internal static void FinishTimeline(Plot plot, Dictionary<double, string> labels)
        {
            TimeAxis(plot, "MM-dd HH:mm");
            plot.Axes.Left.SetTicks(labels.Keys.ToArray(), labels.Values.ToArray());
            plot.Axes.AutoScale();
            plot.Axes.SetLimitsY(-0.5, Math.Max(0.5, labels.Count - 0.5));
            if (labels.Count == 0) EmptyMessage(plot, "No history in this period");
        }

        internal static void EmptyMessage(Plot plot, string message)
        {
            var text = plot.Add.Annotation(message, Alignment.MiddleCenter);
            text.LabelBackgroundColor = Colors.Transparent;
            text.LabelBorderWidth = 0;
            text.LabelShadowColor = Colors.Transparent;
            States.GetOrCreateValue(plot).ApplyColors.Add(theme => text.LabelFontColor = theme.Foreground);
        }

        internal static Plot BackupSizes(List<KeyValuePair<string, double>> values)
        {
            var plot = Create("Backup sizes (GB)");
            plot.HideAxesAndGrid();
            plot.Title("Backup sizes (GB)", 14);
            var slices = values.Where(x => x.Value > 0).Select((x, i) =>
                new PieSlice { Value = x.Value, LegendText = x.Key, Label = x.Key + "\n" + x.Value.ToString("N2") + " GB" }).ToList();
            if (slices.Count == 0) EmptyMessage(plot, "No backups in this period");
            else
            {
                var pie = plot.Add.Pie(slices);
                pie.SliceLabelDistance = 1.45;
                States.GetOrCreateValue(plot).ApplyColors.Add(theme =>
                {
                    pie.LineColor = theme.Foreground;
                    for (int i = 0; i < slices.Count; i++)
                    {
                        slices[i].FillColor = theme.HighContrast ? theme.Background : theme.GetColor(i);
                        slices[i].Fill.Hatch = theme.HighContrast ? Hatch(i) : null;
                        slices[i].Fill.HatchColor = theme.Foreground;
                        slices[i].LabelFontColor = theme.Foreground;
                    }
                });
                plot.Axes.AutoScale();
            }
            SetHover(plot, point => string.Join("\n", values.Select(x => x.Key + ": " + x.Value.ToString("N2") + " GB")));
            return plot;
        }
    }
}
