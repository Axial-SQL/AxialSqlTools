using Microsoft.Data.SqlClient;
using System;
using System.Data;
using System.Threading;

namespace AxialSqlTools.DataCompare
{
    public static class SqlSynchronizer
    {
        public static long Apply(string connectionString, ComparisonResult result, CancellationToken token,
            IProgress<ComparisonProgress> progress = null)
        {
            SynchronizationScript.Validate(result);
            // Validate every selected action before starting a write transaction.
            foreach (var row in result.SelectedRows(token)) SynchronizationScript.RowStatement(result.Plan, row);
            using (var connection = SqlTableReader.Open(connectionString, token))
            using (var transaction = connection.BeginTransaction(IsolationLevel.Serializable))
            {
                bool identity = SynchronizationScript.NeedsIdentityInsert(result);
                try
                {
                    Execute(connection, transaction, "SET XACT_ABORT ON;\n" + SynchronizationScript.DestinationGuard(result.Plan) + SynchronizationScript.SchemaGuard(result.Plan), result, token);
                    if (identity) Execute(connection, transaction, "SET IDENTITY_INSERT " + result.Plan.Target.QualifiedName + " ON;", result, token);
                    long count = 0;
                    foreach (var row in result.SelectedRows(token))
                    {
                        Execute(connection, transaction, SynchronizationScript.RowStatement(result.Plan, row), result, token);
                        if (++count % 100 == 0) progress?.Report(new ComparisonProgress { Phase = "Applying (uncommitted)", Rows = count });
                    }
                    if (identity) Execute(connection, transaction, "SET IDENTITY_INSERT " + result.Plan.Target.QualifiedName + " OFF;", result, token);
                    token.ThrowIfCancellationRequested();
                    transaction.Commit();
                    return count;
                }
                catch
                {
                    try { transaction.Rollback(); } catch (InvalidOperationException) { } catch (SqlException) { }
                    // Do not return a cancelled or failed session with IDENTITY_INSERT set to the pool.
                    SqlConnection.ClearPool(connection);
                    throw;
                }
            }
        }

        private static void Execute(SqlConnection connection, SqlTransaction transaction, string sql, ComparisonResult result, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            using (var command = new SqlCommand(sql, connection, transaction) { CommandTimeout = result.Plan.Options.CommandTimeoutSeconds })
            using (SqlTableReader.CancelWith(command, token)) command.ExecuteNonQuery();
        }
    }
}
