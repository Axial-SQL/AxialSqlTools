using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace AxialSqlTools
{


    public class GitHubReleaseChecker
    {
        private readonly string _owner;
        private readonly string _repository;

        public GitHubReleaseChecker()
        {
            _owner = "Axial-SQL";
            _repository = "AxialSqlTools";
        }

        public async Task<bool> IsNewVersionAvailableAsync(string currentVersion)
        {
            using (var client = new HttpClient())
            {
                // GitHub API versioning
                client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("Axial-SQL-Tools", "Latest"));
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github.v3+json"));

                // Request the latest release from GitHub API
                var url = $"https://api.github.com/repos/{_owner}/{_repository}/releases/latest";
                var response = await client.GetStringAsync(url);

                dynamic latestRelease = JsonConvert.DeserializeObject(response);
                var latestVersion = (string)latestRelease.tag_name;

                // Compare versions (this might need to be more sophisticated depending on your versioning scheme)
                return Version.Parse(latestVersion) > Version.Parse(currentVersion);
            }
        }
    }

}