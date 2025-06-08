using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Smo;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Collections.Specialized;

namespace AxialSqlTools
{
    /// <summary>
    /// Interaction logic for DatabaseScripterToolWindowControl.
    /// </summary>
    public partial class DatabaseScripterToolWindowControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DatabaseScripterToolWindowControl"/> class.
        /// </summary>
        public DatabaseScripterToolWindowControl()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Handles click on the button by displaying a message box.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]
        private void button1_Click(object sender, RoutedEventArgs e)
        {

            var ci = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();

            var sourceConnectionString = ci.FullConnectionString;

            var serverConnection = new ServerConnection(ci.ServerName);
            var server = new Server(serverConnection);

            var chosenDb = server.Databases[ci.Database];   

            string outputRoot = @"C:\temp";
            string dbFolderPath = Path.Combine(outputRoot, chosenDb.Name);

            //
            // 4) Create subdirectories for each object type we intend to script.
            //
            var objectFolders = new[]
            {
                "Tables",
                "Views",
                "StoredProcedures",
                "UserDefinedFunctions",
                "ScalarFunctions",
                "TableValuedFunctions",
                "Synonyms",
                "Types",
                "Triggers",
                "ExtendedProperties"
                // add more folders if you need other object categories
            };

            foreach (var folder in objectFolders)
            {
                string fullPath = Path.Combine(dbFolderPath, folder);
                if (!Directory.Exists(fullPath))
                {
                    Directory.CreateDirectory(fullPath);
                }
            }

            //
            // 5) Prepare common ScriptingOptions (no header comments, include schema, DRI, etc.)
            //
            var options = new ScriptingOptions
            {
                IncludeHeaders = false,  // <--- REMOVE "Script Date", "Server Version", etc.
                IncludeIfNotExists = false,
                ScriptSchema = true,
                DriAll = true,   // script out keys, indexes, FKs, etc.
                Indexes = true,
                SchemaQualify = true,
                NoCollation = false,
                AnsiFile = false,
                ScriptData = false, // only schema, no data
                ScriptBatchTerminator = true,   // include "GO"
                IncludeDatabaseContext = false,  // do NOT include "USE [dbname]"
            };

            //
            // Helper function to write out an SMO UrnCollection (script lines) to a file.
            //
            void WriteScriptToFile(UrnCollection urns, string targetPath, bool formatCode = false)
            {
                // The Scripter will produce a StringCollection of lines.
                var scripter = new Scripter(server)
                {
                    Options = options
                };

                var scriptLines = scripter.Script(urns);
                // Combine lines into one string with Environment.NewLine
                string singleScript = string.Join(Environment.NewLine, scriptLines.Cast<string>().ToArray());

                if (formatCode)
                {                   
                    singleScript = TSqlFormatter.FormatCode(singleScript);
                }   

                File.WriteAllText(targetPath, singleScript);
            }

            //
            // 6a) Script Tables
            //
            foreach (Table table in chosenDb.Tables)
            {
                if (table.IsSystemObject)
                    continue;

                // URN for this table
                var urns = new UrnCollection { table.Urn };
                string outFile = Path.Combine(dbFolderPath, "Tables", $"{table.Schema}.{table.Name}.sql");
                WriteScriptToFile(urns, outFile, formatCode: true);
                Console.WriteLine($"[Table]           {table.Schema}.{table.Name}  →  {outFile}");
            }

            //
            // 6b) Script Views
            //
            foreach (View view in chosenDb.Views)
            {
                if (view.IsSystemObject)
                    continue;

                var urns = new UrnCollection { view.Urn };
                string outFile = Path.Combine(dbFolderPath, "Views", $"{view.Schema}.{view.Name}.sql");
                WriteScriptToFile(urns, outFile);
                Console.WriteLine($"[View]            {view.Schema}.{view.Name}  →  {outFile}");
            }

            //
            // 6c) Script Stored Procedures
            //
            foreach (StoredProcedure sp in chosenDb.StoredProcedures)
            {
                if (sp.IsSystemObject)
                    continue;

                var urns = new UrnCollection { sp.Urn };
                string outFile = Path.Combine(dbFolderPath, "StoredProcedures", $"{sp.Schema}.{sp.Name}.sql");
                WriteScriptToFile(urns, outFile);
                Console.WriteLine($"[StoredProc]      {sp.Schema}.{sp.Name}  →  {outFile}");
            }

            //
            // 6d) Script Scalar and Table-Valued Functions
            //
            foreach (UserDefinedFunction fn in chosenDb.UserDefinedFunctions)
            {
                if (fn.IsSystemObject)
                    continue;

                // Distinguish scalar vs table-valued
                //TODO string folderName = fn.IsTableFunction ? "TableValuedFunctions" : "ScalarFunctions";
                string folderName = "TableValuedFunctions";
                var urns = new UrnCollection { fn.Urn };
                string outFile = Path.Combine(dbFolderPath, folderName, $"{fn.Schema}.{fn.Name}.sql");
                WriteScriptToFile(urns, outFile);
                Console.WriteLine($"[Function]        {fn.Schema}.{fn.Name}  →  {outFile}");
            }

            //
            // 6e) Script Synonyms
            //
            foreach (Synonym syn in chosenDb.Synonyms)
            {
                //if (syn.IsSystemObject)
                //    continue;

                var urns = new UrnCollection { syn.Urn };
                string outFile = Path.Combine(dbFolderPath, "Synonyms", $"{syn.Schema}.{syn.Name}.sql");
                WriteScriptToFile(urns, outFile);
                Console.WriteLine($"[Synonym]         {syn.Schema}.{syn.Name}  →  {outFile}");
            }

            //
            // 6f) Script User-Defined Types
            //
            foreach (UserDefinedDataType udt in chosenDb.UserDefinedDataTypes)
            {
                //if (udt.IsSystemObject)
                //    continue;

                var urns = new UrnCollection { udt.Urn };
                string outFile = Path.Combine(dbFolderPath, "Types", $"{udt.Schema}.{udt.Name}.sql");
                WriteScriptToFile(urns, outFile);
                Console.WriteLine($"[UserType]        {udt.Schema}.{udt.Name}  →  {outFile}");
            }

            //
            // 6g) Script DDL Triggers (on the database)
            //
            foreach (DatabaseDdlTrigger trig in chosenDb.Triggers)
            {
                if (trig.IsSystemObject)
                    continue;

                var urns = new UrnCollection { trig.Urn };
                string outFile = Path.Combine(dbFolderPath, "Triggers", $"{trig.Name}.sql");
                WriteScriptToFile(urns, outFile);
                Console.WriteLine($"[DBTrigger]       {trig.Name}  →  {outFile}");
            }

            ////
            //// 6h) Script Extended Properties (optional, e.g. descriptions)
            ////
            //foreach (ExtendedProperty ep in chosenDb.ExtendedProperties)
            //{
            //    var urns = new UrnCollection { ep.Urn };
            //    string outFile = Path.Combine(dbFolderPath, "ExtendedProperties", $"{ep.Name}.sql");
            //    WriteScriptToFile(urns, outFile);
            //    Console.WriteLine($"[ExtProperty]     {ep.Name}  →  {outFile}");
            //}





        }
    }
}