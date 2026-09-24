using Microsoft.VisualStudio.PlatformUI;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace AxialSqlTools
{
    public partial class SsmsFormatterSettingsDialog : DialogWindow, IDisposable
    {
        private readonly ToolWindowThemeController theme;
        private readonly CancellationTokenSource lifetime = new CancellationTokenSource();
        private SsmsFormatterSettingsSnapshot snapshot;
        private bool loading;
        private bool disposed;

        internal SsmsFormatterSettingsDialog(bool disregardSsmsSettings)
        {
            InitializeComponent();
            theme = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
            ScopeInfo.Text = SsmsFormatterSettingsSnapshot.Scope;
            ModeInfo.Text = "Disregard built-in SSMS Formatter settings (Code Format tab): "
                + (disregardSsmsSettings ? "On. These values are ignored by Axial formatting."
                    : "Off. These values are used before applying Axial options.");
            // Keep the dialog usable on small/high-DPI desktops; all values remain available by scrolling/copying.
            Width = Math.Min(Width, Math.Max(MinWidth, SystemParameters.WorkArea.Width - 40));
            Height = Math.Min(Height, Math.Max(MinHeight, SystemParameters.WorkArea.Height - 40));
            Loaded += Dialog_Loaded;
        }

        private async void Dialog_Loaded(object sender, RoutedEventArgs e)
        {
            Loaded -= Dialog_Loaded;
            await LoadSettingsAsync();
        }

        private async void Refresh_Click(object sender, RoutedEventArgs e)
        {
            await LoadSettingsAsync();
        }

        private async Task LoadSettingsAsync()
        {
            if (loading || disposed) return;
            loading = true;
            RefreshButton.IsEnabled = false;
            CopyButton.IsEnabled = false;
            snapshot = null;
            LeftSettings.ItemsSource = null;
            RightSettings.ItemsSource = null;
            VersionInfo.Text = string.Empty;
            Status.Text = "Loading current SSMS settings...";
            var token = lifetime.Token;
            try
            {
                var result = await SsmsFormatterHost.ReadSettingsAsync(token);
                if (disposed) return;
                token.ThrowIfCancellationRequested();
                snapshot = result;
                VersionInfo.Text = result.Versions;
                int split = (result.Settings.Count + 1) / 2;
                LeftSettings.ItemsSource = result.Settings.Take(split).ToArray();
                RightSettings.ItemsSource = result.Settings.Skip(split).ToArray();
                CopyButton.IsEnabled = true;
                Status.Text = result.Settings.Count + " entries. Take a screenshot or copy all values.";
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested) { }
            catch (Exception ex)
            {
                if (!disposed)
                {
                    AxialSqlToolsPackage._logger?.Warn("Unable to read SSMS formatter settings ({0}).", ex.GetType().Name);
                    Status.Text = "Unable to read SSMS formatter settings. " + ex.Message
                        + " Check that this SSMS version includes SQL Formatter, then click Refresh.";
                }
            }
            finally
            {
                loading = false;
                if (!disposed) RefreshButton.IsEnabled = true;
            }
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            if (snapshot == null || disposed) return;
            try
            {
                Clipboard.SetText(snapshot.ToText(ModeInfo.Text));
                Status.Text = "Copied all " + snapshot.Settings.Count + " entries and version information.";
            }
            catch (ExternalException)
            {
                Status.Text = "The clipboard is busy. Try Copy all again.";
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            Dispose();
            base.OnClosed(e);
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            Loaded -= Dialog_Loaded;
            lifetime.Cancel();
            lifetime.Dispose();
            theme.Dispose();
        }
    }
}
