using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using System;
using System.Text.RegularExpressions;

namespace AxialSqlTools.JobQuickView
{
    // Capture the identity and authentication before the menu closes or focus changes.
    // Never keep an Object Explorer node or a live SSMS connection in the window.
    internal sealed class JobQuickViewSelection
    {
        private readonly SqlConnectionInfo connection;

        private JobQuickViewSelection(SqlConnectionInfo connection, string jobName)
        {
            this.connection = connection.Copy();
            this.connection.DatabaseName = "msdb";
            ServerName = connection.ServerName;
            JobName = jobName;
        }

        public string ServerName { get; }
        public string JobName { get; }

        public static bool IsJob(INodeInformation node)
        {
            if (node == null || !(node.Connection is SqlConnectionInfo)) return false;
            // Folders and JobStep descendants must never be treated as a job.
            return string.Equals(node.UrnPath, "Server/JobServer/Job", StringComparison.OrdinalIgnoreCase);
        }

        public static JobQuickViewSelection Capture(INodeInformation node)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!IsJob(node))
                throw new InvalidOperationException("Select one SQL Server Agent job in Object Explorer, then choose Quick View.");

            // URN attributes preserve apostrophes, brackets, slashes and exact catalog
            // casing. The tree caption may contain localized status decorations.
            string jobName = NameFromUrn(node.NavigationContext) ?? NameFromUrn(node.Context);
            if (string.IsNullOrEmpty(jobName)) jobName = node.InvariantName;
            if (string.IsNullOrEmpty(jobName))
                throw new InvalidOperationException("The selected job name is unavailable. Refresh the Jobs folder and try again.");

            return new JobQuickViewSelection((SqlConnectionInfo)node.Connection, jobName);
        }

        internal static string NameFromUrn(string urn)
        {
            if (string.IsNullOrEmpty(urn)) return null;
            var job = Regex.Match(urn, @"(?:^|/)Job\[(?<attributes>(?:[^\]']|'(?:''|[^'])*')*)\]$",
                RegexOptions.CultureInvariant);
            if (!job.Success) return null;
            foreach (Match attribute in Regex.Matches(job.Groups["attributes"].Value,
                @"@(?<key>\w+)='(?<value>(?:''|[^'])*)'", RegexOptions.CultureInvariant))
            {
                if (attribute.Groups["key"].Value == "Name")
                    return attribute.Groups["value"].Value.Replace("''", "'");
            }
            return null;
        }

        // Called by the data service off the UI thread for every operation. Keep the
        // original SSMS authentication, TLS and network options, including access-token
        // support, instead of reconstructing a connection string from username/password.
        public SqlConnection CreateConnection()
        {
            var created = connection.Copy().CreateConnectionObject();
            if (created is SqlConnection sqlConnection) return sqlConnection;
            created?.Dispose();
            throw new NotSupportedException("Quick View requires the Microsoft.Data.SqlClient connection provider included with SSMS 22.");
        }
    }
}
