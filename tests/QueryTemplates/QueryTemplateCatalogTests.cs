using AxialSqlTools;
using System;
using System.IO;
using System.Linq;
using System.Threading;

internal static class QueryTemplateCatalogTests
{
    private static int _passed;

    private static int Main()
    {
        try
        {
            Run("Root and nested files are sorted and all extensions are retained", SortedTree);
            Run("Existing command identities survive additions and deletions", StableNames);
            Run("Renames get a new identity without changing sibling identities", Rename);
            Run("Command identities normalize path case and trailing separators", NormalizedNames);
            Run("An unreadable descendant does not hide root or sibling scripts", UnreadableChild);
            Run("Unreadable root propagates failure instead of returning an empty catalog", UnreadableRoot);
            Run("Missing root propagates failure", MissingRoot);
            Run("Directory reparse points are skipped without traversing them", ReparsePoint);
            Run("Cancellation before and during discovery propagates", Cancellation);
            Run("Empty folder produces an empty catalog", EmptyFolder);
            Console.WriteLine("Passed " + _passed + " Query Templates catalog tests.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static void Run(string name, Action action)
    {
        action();
        _passed++;
        Console.WriteLine("PASS: " + name);
    }

    private static void SortedTree()
    {
        WithFolder(root =>
        {
            Write(root, "z.sql");
            Write(root, "A.SQL");
            Write(root, "notes.txt");
            Write(root, Path.Combine("z-folder", "child.sql"));
            Write(root, Path.Combine("a-folder", "nested", "child.sql"));
            var catalog = Load(root);
            Equal("a-folder,z-folder", string.Join(",", catalog.Root.Folders.Select(f => f.Name)));
            Equal("A.SQL,notes.txt,z.sql", string.Join(",", catalog.Root.Files.Select(f => f.Name)));
            Equal(5, catalog.FileCount);
            Equal(0, catalog.Warnings.Count);
            Equal(Path.Combine(root, "a-folder", "nested", "child.sql"),
                catalog.Root.Folders[0].Folders[0].Files[0].FullPath);
        });
    }

    private static void StableNames()
    {
        WithFolder(root =>
        {
            Write(root, "root.sql");
            Write(root, Path.Combine("nested", "script.sql"));
            var before = Load(root);
            Write(root, "aaa.sql");
            Write(root, Path.Combine("aaa-folder", "aaa.sql"));
            var added = Load(root);
            Equal(before.Root.Files[0].CommandName, added.Root.Files.Single(f => f.Name == "root.sql").CommandName);
            Equal(before.Root.Folders[0].CommandName, added.Root.Folders.Single(f => f.Name == "nested").CommandName);
            Equal(before.Root.Folders[0].Files[0].CommandName,
                added.Root.Folders.Single(f => f.Name == "nested").Files[0].CommandName);
            File.Delete(Path.Combine(root, "aaa.sql"));
            Directory.Delete(Path.Combine(root, "aaa-folder"), true);
            var deleted = Load(root);
            Equal(before.Root.Files[0].CommandName, deleted.Root.Files[0].CommandName);
            Equal(2, deleted.FileCount);
            Equal(2, before.FileCount); // A later refresh cannot mutate the old snapshot.
        });
    }

    private static void Rename()
    {
        WithFolder(root =>
        {
            Write(root, "before.sql");
            Write(root, "sibling.sql");
            var before = Load(root);
            File.Move(Path.Combine(root, "before.sql"), Path.Combine(root, "after.sql"));
            var after = Load(root);
            Assert(before.Root.Files[0].CommandName != after.Root.Files[0].CommandName, "Renamed file retained old command name.");
            Equal(before.Root.Files[1].CommandName, after.Root.Files[1].CommandName);
            Equal("after.sql,sibling.sql", string.Join(",", after.Root.Files.Select(f => f.Name)));
        });
    }

    private static void NormalizedNames()
    {
        WithFolder(root =>
        {
            var path = Path.Combine(root, "MixedName.SQL");
            Equal(QueryTemplateCatalog.GetCommandName(path, false), QueryTemplateCatalog.GetCommandName(path.ToUpperInvariant(), false));
            Equal(QueryTemplateCatalog.GetCommandName(root, true), QueryTemplateCatalog.GetCommandName(root + Path.DirectorySeparatorChar, true));
            Assert(QueryTemplateCatalog.GetCommandName(path, false) != QueryTemplateCatalog.GetCommandName(path, true),
                "File and folder command names must use different namespaces.");
            Equal(Load(root).Root.FullPath, Load(root + Path.DirectorySeparatorChar).Root.FullPath);
        });
    }

    private static void UnreadableChild()
    {
        WithFolder(root =>
        {
            Write(root, "root.sql");
            Write(root, Path.Combine("bad", "hidden.sql"));
            Write(root, Path.Combine("good", "visible.sql"));
            string denied = Path.Combine(root, "bad");
            var fileSystem = new TestFileSystem { FailFilesPath = denied };
            var catalog = QueryTemplateCatalog.Load(root, CancellationToken.None, fileSystem);
            Equal(2, catalog.FileCount);
            Equal("root.sql", catalog.Root.Files[0].Name);
            Equal("good", catalog.Root.Folders.Single().Name);
            Equal(1, catalog.Warnings.Count);
            Assert(catalog.Warnings[0].Contains(denied), "Warning did not identify the inaccessible folder.");
        });
    }

    private static void UnreadableRoot()
    {
        WithFolder(root =>
        {
            Write(root, "root.sql");
            Throws<UnauthorizedAccessException>(() => QueryTemplateCatalog.Load(root, CancellationToken.None,
                new TestFileSystem { FailFilesPath = root }));
        });
    }

    private static void MissingRoot()
    {
        WithFolder(root => Throws<DirectoryNotFoundException>(() => Load(Path.Combine(root, "missing"))));
    }

    private static void ReparsePoint()
    {
        WithFolder(root =>
        {
            Write(root, "root.sql");
            Write(root, Path.Combine("link", "never-read.sql"));
            string link = Path.Combine(root, "link");
            var fileSystem = new TestFileSystem { ReparsePath = link, FailFilesPath = link };
            var catalog = QueryTemplateCatalog.Load(root, CancellationToken.None, fileSystem);
            Equal(1, catalog.FileCount);
            Equal(0, catalog.Root.Folders.Count);
            Equal(1, catalog.Warnings.Count);
            Assert(catalog.Warnings[0].Contains("Skipped linked"), "Linked folder was traversed.");
        });
    }

    private static void Cancellation()
    {
        WithFolder(root =>
        {
            using (var before = new CancellationTokenSource())
            {
                before.Cancel();
                Throws<OperationCanceledException>(() => QueryTemplateCatalog.Load(root, before.Token));
            }
            using (var empty = new CancellationTokenSource())
            {
                var fileSystem = new TestFileSystem { BeforeGetFiles = _ => empty.Cancel() };
                Throws<OperationCanceledException>(() => QueryTemplateCatalog.Load(root, empty.Token, fileSystem));
            }
            Write(root, "root.sql");
            Write(root, Path.Combine("nested", "child.sql"));
            using (var during = new CancellationTokenSource())
            {
                var fileSystem = new TestFileSystem { BeforeGetFiles = _ => during.Cancel() };
                Throws<OperationCanceledException>(() => QueryTemplateCatalog.Load(root, during.Token, fileSystem));
            }
        });
    }

    private static void EmptyFolder()
    {
        WithFolder(root =>
        {
            var catalog = Load(root);
            Equal(0, catalog.FileCount);
            Equal(0, catalog.Root.Files.Count);
            Equal(0, catalog.Root.Folders.Count);
            Equal(0, catalog.Warnings.Count);
        });
    }

    private static QueryTemplateCatalog Load(string root) => QueryTemplateCatalog.Load(root, CancellationToken.None);

    private static void Write(string root, string relativePath)
    {
        string path = Path.Combine(root, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(path));
        File.WriteAllText(path, "select 1;");
    }

    private static void WithFolder(Action<string> test)
    {
        string root = Path.Combine(Path.GetTempPath(), "AxialSqlTools-templates-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try { test(root); }
        finally { Directory.Delete(root, true); }
    }

    private static void Equal<T>(T expected, T actual)
    {
        Assert(object.Equals(expected, actual), "Expected '" + expected + "', got '" + actual + "'.");
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    private static void Throws<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new InvalidOperationException("Expected " + typeof(T).Name + ".");
    }

    private sealed class TestFileSystem : IQueryTemplateFileSystem
    {
        public string FailFilesPath;
        public string ReparsePath;
        public Action<string> BeforeGetFiles;

        public string[] GetDirectories(string path) => Directory.GetDirectories(path);

        public string[] GetFiles(string path)
        {
            BeforeGetFiles?.Invoke(path);
            if (path == FailFilesPath) throw new UnauthorizedAccessException("Simulated access denied.");
            return Directory.GetFiles(path);
        }

        public FileAttributes GetAttributes(string path)
        {
            return path == ReparsePath ? FileAttributes.Directory | FileAttributes.ReparsePoint : File.GetAttributes(path);
        }
    }
}
