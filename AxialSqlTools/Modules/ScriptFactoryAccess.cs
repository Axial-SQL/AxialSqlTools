using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

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

            //I assume there is no harm in always encrypting these connections
            //SSMS 20 connection fails without this
            builder.Encrypt = true; // TODO - take from advance options
            builder.TrustServerCertificate = true; // TODO - take from advance options

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
