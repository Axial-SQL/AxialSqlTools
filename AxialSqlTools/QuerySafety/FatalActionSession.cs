using System;
using System.Security.Cryptography;
using System.Text;

namespace AxialSqlTools.QuerySafety
{
    // One instance per editor document, released on closing. Never persisted to settings.
    internal sealed class FatalActionSession
    {
        private string approvedSql;
        private string approvedConnection;

        internal bool IsApproved(string sql, string connection)
        {
            bool matches = connection != null && approvedConnection == connection && approvedSql == Fingerprint(sql);
            if (!matches) Clear(); // Edited SQL, reconnects or target changes end the approval.
            return matches;
        }

        internal void Approve(string sql, string connection)
        {
            if (connection == null) return;
            approvedSql = Fingerprint(sql);
            approvedConnection = connection;
        }

        internal void Clear() { approvedSql = null; approvedConnection = null; }

        private static string Fingerprint(string sql)
        {
            using (var hash = SHA256.Create())
                return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(sql ?? "")));
        }
    }
}
