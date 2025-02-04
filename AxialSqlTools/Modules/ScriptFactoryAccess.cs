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

            Match match = Regex.Match(selectedNode.Context, @"Database\[@Name='(.*?)'\]");
            if (match.Success)
            {
                databaseName = match.Groups[1].Value;
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

            UIConnectionInfo connection = ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo;

            string databaseName = inMaster ? "master" : connection.AdvancedOptions["DATABASE"];
            if (string.IsNullOrEmpty(databaseName))
                databaseName = "master";

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

            builder.DataSource = connection.ServerName;
            builder.IntegratedSecurity = string.IsNullOrEmpty(connection.Password);
            builder.Password = connection.Password;
            builder.UserID = connection.UserName;
            builder.InitialCatalog = databaseName;
            builder.ApplicationName = "Axial SQL Tools";

            if (connection.AdvancedOptions["ENCRYPT_CONNECTION"] == "True")          
            {
                builder.Encrypt = true;
            }
            if (connection.AdvancedOptions["TRUST_SERVER_CERTIFICATE"] == "True")
            {
                builder.TrustServerCertificate = true;
            }

            ConnectionInfo ci = new ConnectionInfo();
            ci.FullConnectionString = builder.ToString();
            ci.Database = databaseName;
            ci.ServerName = connection.ServerName;
            ci.ActiveConnectionInfo = connection;

            return ci;

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
