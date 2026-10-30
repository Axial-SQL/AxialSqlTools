using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace AxialSqlTools
{
    // A filesystem snapshot, independent of the SSMS command bars. Read it before
    // replacing a working menu so an unavailable templates root does not erase it.
    internal sealed class QueryTemplateCatalog
    {
        private QueryTemplateCatalog(QueryTemplateFolder root, List<string> warnings)
        {
            Root = root;
            Warnings = warnings.AsReadOnly();
            FileCount = CountFiles(root);
        }

        public QueryTemplateFolder Root { get; }
        public IReadOnlyList<string> Warnings { get; }
        public int FileCount { get; }

        public static QueryTemplateCatalog Load(string folderPath, CancellationToken cancellationToken)
        {
            return Load(folderPath, cancellationToken, new QueryTemplateFileSystem());
        }

        internal static QueryTemplateCatalog Load(string folderPath, CancellationToken cancellationToken,
            IQueryTemplateFileSystem fileSystem)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("A templates folder must be configured.", nameof(folderPath));
            if (fileSystem == null)
                throw new ArgumentNullException(nameof(fileSystem));

            cancellationToken.ThrowIfCancellationRequested();
            string fullPath = Path.GetFullPath(folderPath);
            string rootPath = Path.GetPathRoot(fullPath);
            if (fullPath.Length > rootPath.Length)
                fullPath = fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            var warnings = new List<string>();
            // Do not catch failure here: the caller must retain the old menu when
            // the configured root itself could not be read.
            var root = ReadFolder(fullPath, cancellationToken, fileSystem, warnings);
            return new QueryTemplateCatalog(root, warnings);
        }

        private static QueryTemplateFolder ReadFolder(string path, CancellationToken cancellationToken,
            IQueryTemplateFileSystem fileSystem, List<string> warnings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string[] directories = fileSystem.GetDirectories(path);
            string[] files = fileSystem.GetFiles(path);
            cancellationToken.ThrowIfCancellationRequested();
            Array.Sort(directories, ComparePaths);
            Array.Sort(files, ComparePaths);

            var templates = new List<QueryTemplateFile>();
            foreach (string file in files)
            {
                cancellationToken.ThrowIfCancellationRequested();
                // Keep the previous behavior of including every file extension.
                templates.Add(new QueryTemplateFile(file));
            }

            var folders = new List<QueryTemplateFolder>();
            foreach (string directory in directories)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    // Junctions and symlinks can lead back to an ancestor or out
                    // to an unexpectedly large tree. Do not recurse through them.
                    if ((fileSystem.GetAttributes(directory) & FileAttributes.ReparsePoint) != 0)
                    {
                        warnings.Add("Skipped linked templates folder: " + directory);
                        continue;
                    }
                    folders.Add(ReadFolder(directory, cancellationToken, fileSystem, warnings));
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is SecurityException)
                {
                    // One bad descendant must not prevent root-level scripts or
                    // sibling folders from appearing.
                    warnings.Add("Could not read templates folder '" + directory + "': " + ex.Message);
                }
            }

            return new QueryTemplateFolder(path, folders, templates);
        }

        private static int ComparePaths(string left, string right)
        {
            int comparison = StringComparer.OrdinalIgnoreCase.Compare(Path.GetFileName(left), Path.GetFileName(right));
            return comparison != 0 ? comparison : StringComparer.Ordinal.Compare(left, right);
        }

        private static int CountFiles(QueryTemplateFolder folder)
        {
            int count = folder.Files.Count;
            foreach (QueryTemplateFolder child in folder.Folders)
                count += CountFiles(child);
            return count;
        }

        internal static string GetCommandName(string fullPath, bool isFolder)
        {
            // Windows paths are case insensitive. Naming by path keeps command
            // identities stable when another file is added, removed, or reordered.
            string normalized = Path.GetFullPath(fullPath).Replace('\\', '/').TrimEnd('/').ToUpperInvariant();
            using (var sha = SHA256.Create())
            {
                string hash = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(normalized))).Replace("-", string.Empty);
                return (isFolder ? "AxialSqlTools_Folder_" : "AxialSqlTools_Template_") + hash;
            }
        }
    }

    internal sealed class QueryTemplateFolder
    {
        public QueryTemplateFolder(string fullPath, List<QueryTemplateFolder> folders, List<QueryTemplateFile> files)
        {
            FullPath = fullPath;
            Name = Path.GetFileName(fullPath);
            CommandName = QueryTemplateCatalog.GetCommandName(fullPath, true);
            Folders = folders.AsReadOnly();
            Files = files.AsReadOnly();
        }

        public string FullPath { get; }
        public string Name { get; }
        public string CommandName { get; }
        public IReadOnlyList<QueryTemplateFolder> Folders { get; }
        public IReadOnlyList<QueryTemplateFile> Files { get; }
    }

    internal sealed class QueryTemplateFile
    {
        public QueryTemplateFile(string fullPath)
        {
            FullPath = fullPath;
            Name = Path.GetFileName(fullPath);
            CommandName = QueryTemplateCatalog.GetCommandName(fullPath, false);
        }

        public string FullPath { get; }
        public string Name { get; }
        public string CommandName { get; }
    }

    internal interface IQueryTemplateFileSystem
    {
        string[] GetDirectories(string path);
        string[] GetFiles(string path);
        FileAttributes GetAttributes(string path);
    }

    internal sealed class QueryTemplateFileSystem : IQueryTemplateFileSystem
    {
        public string[] GetDirectories(string path) => Directory.GetDirectories(path);
        public string[] GetFiles(string path) => Directory.GetFiles(path);
        public FileAttributes GetAttributes(string path) => File.GetAttributes(path);
    }
}
