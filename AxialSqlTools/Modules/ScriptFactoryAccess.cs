using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

            ConnectionInfo ci = new ConnectionInfo();
            ci.FullConnectionString = builder.ToString();
            ci.Database = databaseName;
            ci.ServerName = connection.ServerName;
            ci.ActiveConnectionInfo = connection;

            return ci;

        }

    }
}
