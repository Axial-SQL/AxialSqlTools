using Microsoft.Data.SqlClient;
using System;
using System.Windows;
using System.Windows.Controls;

namespace AxialSqlTools.DataCompare
{
    public partial class SqlConnectionDialog : Microsoft.VisualStudio.PlatformUI.DialogWindow
    {
        private readonly ToolWindowThemeController theme;
        public string ConnectionString { get; private set; }
        public SqlConnectionDialog(string current)
        {
            InitializeComponent();
            theme = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
            if (!string.IsNullOrEmpty(current))
            {
                var builder = new SqlConnectionStringBuilder(current);
                ServerBox.Text = builder.DataSource; DatabaseBox.Text = builder.InitialCatalog;
                UsernameBox.Text = builder.UserID;
                // Do not copy a password into a new connection dialog.
            }
            Closed += (s, e) => { PasswordBox.Clear(); theme.Dispose(); };
        }

        private void AuthenticationChanged(object sender, SelectionChangedEventArgs e)
        {
            if (UsernameBox == null || PasswordBox == null) return;
            UsernameBox.IsEnabled = AuthenticationBox.SelectedIndex != 0;
            PasswordBox.IsEnabled = AuthenticationBox.SelectedIndex == 1;
            if (!PasswordBox.IsEnabled) PasswordBox.Clear();
        }

        private void ConnectClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ServerBox.Text) || string.IsNullOrWhiteSpace(DatabaseBox.Text))
            { MessageBox.Show(this, "Enter a server and database.", Title); return; }
            var builder = new SqlConnectionStringBuilder { DataSource = ServerBox.Text.Trim(), InitialCatalog = DatabaseBox.Text.Trim(),
                ApplicationName = "Axial SQL Tools - Data Compare", ConnectTimeout = 30, Encrypt = EncryptBox.IsChecked == true,
                TrustServerCertificate = TrustBox.IsChecked == true, PersistSecurityInfo = false };
            if (AuthenticationBox.SelectedIndex == 0) builder.IntegratedSecurity = true;
            else
            {
                if (AuthenticationBox.SelectedIndex == 1)
                {
                    if (string.IsNullOrWhiteSpace(UsernameBox.Text)) { MessageBox.Show(this, "Enter a SQL Server user name.", Title); return; }
                    builder.UserID = UsernameBox.Text.Trim(); builder.Password = PasswordBox.Password;
                }
                else
                {
                    builder.Authentication = AuthenticationBox.SelectedIndex == 2 ? SqlAuthenticationMethod.ActiveDirectoryInteractive : SqlAuthenticationMethod.ActiveDirectoryDefault;
                    if (!string.IsNullOrWhiteSpace(UsernameBox.Text)) builder.UserID = UsernameBox.Text.Trim();
                }
            }
            ConnectionString = builder.ConnectionString;
            DialogResult = true;
        }
    }
}
