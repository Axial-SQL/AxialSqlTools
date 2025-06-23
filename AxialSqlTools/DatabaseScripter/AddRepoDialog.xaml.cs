// AddRepoDialog.xaml.cs
using Octokit;
using System;
using System.Windows;

namespace AxialSqlTools
{
    public partial class AddRepoDialog : Window
    {
        public GitRepo Result { get; private set; }

        public AddRepoDialog()
        {
            InitializeComponent();
        }

        private async void Ok_Click(object sender, RoutedEventArgs e)
        {
            var owner = OwnerBox.Text.Trim();
            var name = NameBox.Text.Trim();
            var branch = BranchBox.Text.Trim();
            var token = TokenBox.Text.Trim();
            if (string.IsNullOrEmpty(owner) || string.IsNullOrEmpty(name) ||
                string.IsNullOrEmpty(branch) || string.IsNullOrEmpty(token))
            {
                MessageBox.Show("All fields are required.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var client = new GitHubClient(new ProductHeaderValue("AxialSqlTools"))
                { Credentials = new Credentials(token) };
                // test access/app existence
                var repo = await client.Repository.Get(owner, name);
                // test branch exists
                await client.Repository.Branch.Get(owner, name, branch);

                Result = new GitRepo
                {
                    Owner = owner,
                    Name = name,
                    Branch = branch,
                    Token = token
                };
                DialogResult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to validate repo/branch/token:\n{ex.Message}",
                                "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
