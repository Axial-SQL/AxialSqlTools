using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.SqlServer.Management.Common;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Text.RegularExpressions;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace AxialSqlTools
{
    public static class ScriptFactoryAccess
    {
        public class ConnectionInfo
        {          
            public string FullConnectionString { get; set; }
            public string Database { get; set; }            
            public string ServerName { get; set; }            
            public UIConnectionInfo ActiveConnectionInfo { get; set; }

            public string DisplayName
            {
                get
                {
                    return $"[{ServerName}] \\ [{Database}]";
                }
            }

            public override string ToString() => DisplayName;
        }

        private static INodeInformation GetSelectedNode(IObjectExplorerService _objectExplorerService)
        {
            INodeInformation[] nodes;
            int nodeCount;
            _objectExplorerService.GetSelectedNodes(out nodeCount, out nodes);

            return (nodeCount > 0 ? nodes[0] : null);
        }

        public static ConnectionInfo GetCurrentConnectionInfoFromObjectExplorer(bool inMaster = false)
        {

            var oeService = (IObjectExplorerService)ServiceCache.ServiceProvider.GetService(typeof(IObjectExplorerService));
            if (oeService == null)
                return null;

            // Get the currently selected nodes in Object Explorer.
            var selectedNode = GetSelectedNode(oeService);

            if (selectedNode == null)
                return null;

            string databaseName = "master";
            if (!inMaster)
            {
                Match match = Regex.Match(selectedNode.Context, @"Database\[@Name='(.*?)'\]");
                if (match.Success)
                {
                    databaseName = match.Groups[1].Value;
                }
            }

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = selectedNode.Connection.ServerName;
            builder.InitialCatalog = databaseName;
            builder.UserID = selectedNode.Connection.UserName;
            builder.IntegratedSecurity = string.IsNullOrEmpty(selectedNode.Connection.Password);
            builder.Password = selectedNode.Connection.Password;
            builder.ApplicationName = "Axial SQL Tools";
            builder.Encrypt = ((SqlConnectionInfo)selectedNode.Connection).EncryptConnection;
            builder.TrustServerCertificate = ((SqlConnectionInfo)selectedNode.Connection).TrustServerCertificate;


            ConnectionInfo ci = new ConnectionInfo();
            ci.FullConnectionString = builder.ToString();
            ci.Database = databaseName;
            ci.ServerName = builder.DataSource;

            return ci;
        }

        public static ConnectionInfo GetCurrentConnectionInfo(bool inMaster = false)
        {
            var scriptFactory = ServiceCache.ScriptFactory;
            if (scriptFactory == null) return null;

            var connInfo = scriptFactory.CurrentlyActiveWndConnectionInfo;
            if (connInfo == null) return null;

            UIConnectionInfo connection = connInfo.UIConnectionInfo;
            if (connection == null) return null;

            string databaseName = inMaster ? "master" : GetAdvancedOption(connection, "DATABASE");
            if (string.IsNullOrWhiteSpace(databaseName))
                databaseName = "master";

            var builder = new SqlConnectionStringBuilder
            {
                DataSource = connection.ServerName,
                InitialCatalog = databaseName,
                ApplicationName = "Axial SQL Tools"
            };

            string auth = GetAuthenticationMode(connection);

            if (ShouldUseMicrosoftEntraInteractive(connection, auth))
            {
                builder.Authentication = SqlAuthenticationMethod.ActiveDirectoryInteractive;

                if (!string.IsNullOrWhiteSpace(connection.UserName))
                    builder.UserID = connection.UserName;

                // Do not set Password for Microsoft Entra MFA.
                // Do not set IntegratedSecurity=True, because that triggers SSPI/Kerberos.
            }
            else if (!string.IsNullOrWhiteSpace(connection.Password))
            {
                builder.IntegratedSecurity = false;
                builder.UserID = connection.UserName;
                builder.Password = connection.Password;
            }
            else
            {
                builder.IntegratedSecurity = true;
            }

            if (IsTrue(GetAdvancedOption(connection, "ENCRYPT_CONNECTION")))
                builder.Encrypt = true;

            if (IsTrue(GetAdvancedOption(connection, "TRUST_SERVER_CERTIFICATE")))
                builder.TrustServerCertificate = true;

            var ci = new ConnectionInfo
            {
                FullConnectionString = builder.ToString(),
                Database = databaseName,
                ServerName = connection.ServerName,
                ActiveConnectionInfo = connection
            };

            return ci;
        }

        private static string GetAdvancedOption(UIConnectionInfo connection, string key)
        {
            if (connection?.AdvancedOptions == null || string.IsNullOrWhiteSpace(key))
                return null;

            return connection.AdvancedOptions.Get(key);
        }

        private static string GetAuthenticationMode(UIConnectionInfo connection)
        {
            if (connection == null)
                return string.Empty;

            var values = new List<string>();

            if (!string.IsNullOrWhiteSpace(connection.OtherParams))
                values.Add(connection.OtherParams);

            if (connection.AdvancedOptions != null)
            {
                foreach (string key in connection.AdvancedOptions.AllKeys)
                {
                    if (!string.IsNullOrWhiteSpace(key))
                        values.Add(key);

                    string value = connection.AdvancedOptions.Get(key);
                    if (!string.IsNullOrWhiteSpace(value))
                        values.Add(value);
                }
            }

            return string.Join(";", values);
        }

        private static bool ShouldUseMicrosoftEntraInteractive(UIConnectionInfo connection, string auth)
        {
            if (IsMicrosoftEntraMfa(auth))
                return true;

            if (connection == null)
                return false;

            bool hasUserName = !string.IsNullOrWhiteSpace(connection.UserName);
            bool hasPassword = !string.IsNullOrWhiteSpace(connection.Password);

            // Microsoft Entra MFA commonly has a UPN username and no password.
            // This prevents accidental fallback to Integrated Security=True / SSPI.
            if (hasUserName && !hasPassword && LooksLikeUpn(connection.UserName))
                return true;

            return false;
        }

        private static bool IsMicrosoftEntraMfa(string auth)
        {
            if (string.IsNullOrWhiteSpace(auth))
                return false;

            return auth.IndexOf("Active Directory Interactive", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   auth.IndexOf("ActiveDirectoryInteractive", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   auth.IndexOf("Microsoft Entra MFA", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   auth.IndexOf("Microsoft Entra", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   auth.IndexOf("Azure Active Directory - Universal with MFA", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   auth.IndexOf("Universal with MFA", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   auth.IndexOf("MFA", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool LooksLikeUpn(string userName)
        {
            if (string.IsNullOrWhiteSpace(userName))
                return false;

            return userName.Contains("@") && !userName.Contains("\\");
        }

        private static bool IsTrue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            return value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("1", StringComparison.OrdinalIgnoreCase);
        }

        public static string GetActiveQueryWindowText()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                DTE application = ServiceCache.ExtensibilityModel?.Application;
                if (application == null || application.ActiveDocument == null)
                {
                    return string.Empty;
                }

                TextDocument doc = application.ActiveDocument.Object("TextDocument") as TextDocument;
                if (doc == null)
                {
                    return string.Empty;
                }

                EditPoint startPoint = doc.StartPoint.CreateEditPoint();
                return startPoint.GetText(doc.EndPoint);
            }
            catch
            {
                return string.Empty;
            }
        }

        public static string GetXmlFromUIConnectionInfo(UIConnectionInfo connectionInfo)
        {
            if (connectionInfo == null)
            {
                throw new ArgumentNullException(nameof(connectionInfo), "ConnectionInfo cannot be null.");
            }

            StringBuilder sb = new StringBuilder();

            using (XmlWriter writer = XmlWriter.Create(sb))
            {
                connectionInfo.SaveToStream(writer, saveName: true);
            }

            return sb.ToString();
        }

        public static UIConnectionInfo CreateConnectionInfoFromXml(string xmlString)
        {
            if (string.IsNullOrEmpty(xmlString))
            {
                throw new ArgumentNullException(nameof(xmlString), "The XML string cannot be null or empty.");
            }

            UIConnectionInfo connectionInfo = null;

            using (var xmlReader = XmlReader.Create(new StringReader(xmlString)))
            {
                xmlReader.MoveToContent();
                connectionInfo = UIConnectionInfo.LoadFromStream(xmlReader);
            }

            return connectionInfo;
        }

    }
}
