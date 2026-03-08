using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        private static bool TryGetConnectionInfo(IVsWindowFrame frame, out string server, out string database)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            server = null;
            database = null;

            if (frame == null)
            {
                return false;
            }

            try
            {
                // 1. DocView/Reflection Method (Main attempt)
                if (TryPopulateFromDocView(frame, out string docViewServer, out string docViewDatabase))
                {
                    server = docViewServer;
                    database = docViewDatabase;
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                EnvTabsLog.Error($"TryGetConnectionInfo exception: {ex.Message}");
                return false;
            }
        }

        private static bool TryPopulateFromDocView(IVsWindowFrame frame, out string server, out string database)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            server = null;
            database = null;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out object docView) == VSConstants.S_OK && docView != null)
                {
                    // Inspect private field "m_connection" on the DocView (SqlScriptEditorControl)
                    // This is more reliable than tooltips which users can disable.
                    var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
                    var field = docView.GetType().GetField("m_connection", flags);
                    if (field != null)
                    {
                        var connection = field.GetValue(docView);
                        if (connection != null)
                        {
                            var connType = connection.GetType();
                            // properties: DataSource (=Server), Database
                            var propDataSource = connType.GetProperty("DataSource");
                            var propDatabase = connType.GetProperty("Database");
                            
                            string s = propDataSource?.GetValue(connection) as string;
                            string d = propDatabase?.GetValue(connection) as string;

                            if (!string.IsNullOrWhiteSpace(s))
                            {
                                server = s;
                                database = d;
                                return true;
                            }
                        }
                    }
                    else
                    {
                        // ignore
                    }
                }
            }
            catch (Exception ex)
            {
                // Fail silently, fallback to other methods
                EnvTabsLog.Error($"Error accessing DocView connection info: {ex.Message}");
            }
            return false;
        }

        private static void TryPopulateFromCaptions(IVsWindowFrame frame, ref string server, ref string database)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (frame == null) return;

            string caption = TryReadFrameCaption(frame);
            if (TryParseServerDatabaseFromCaption(caption, out string s1, out string d1))
            {
                if (string.IsNullOrWhiteSpace(server)) server = s1;
                if (string.IsNullOrWhiteSpace(database)) database = d1;
                return;
            }

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_EditorCaption, out object editorCaptionObj) == VSConstants.S_OK)
                {
                    string editorCaption = editorCaptionObj as string;
                    if (TryParseServerDatabaseFromCaption(editorCaption, out string s2, out string d2))
                    {
                        if (string.IsNullOrWhiteSpace(server)) server = s2;
                        if (string.IsNullOrWhiteSpace(database)) database = d2;
                    }
                }
            }
            catch
            {
                // ignore
            }
        }

        private static bool TryParseServerDatabaseFromCaption(string caption, out string server, out string database)
        {
            server = null;
            database = null;

            if (string.IsNullOrWhiteSpace(caption)) return false;

            try
            {
                int dash = caption.IndexOf(" - ", StringComparison.Ordinal);
                if (dash < 0)
                {
                    return false;
                }

                string tail = caption.Substring(dash + 3);
                int paren = tail.IndexOf(" (", StringComparison.Ordinal);
                if (paren >= 0)
                {
                    tail = tail.Substring(0, paren);
                }

                tail = tail.Trim();
                if (string.IsNullOrWhiteSpace(tail)) return false;

                int dot = tail.IndexOf('.');
                if (dot > 0 && dot < tail.Length - 1)
                {
                    server = tail.Substring(0, dot).Trim();
                    database = tail.Substring(dot + 1).Trim();
                }
                else
                {
                    server = tail;
                }

                return !string.IsNullOrWhiteSpace(server) || !string.IsNullOrWhiteSpace(database);
            }
            catch
            {
                return false;
            }
        }

        private List<OpenDocumentInfo> GetOpenDocumentsSnapshot()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var list = new List<OpenDocumentInfo>();

            rdt.GetRunningDocumentsEnum(out IEnumRunningDocuments enumDocs);
            if (enumDocs == null)
            {
                return list;
            }

            uint[] cookies = new uint[1];
            while (enumDocs.Next(1, cookies, out uint fetched) == VSConstants.S_OK && fetched == 1)
            {
                uint cookie = cookies[0];
                string moniker = TryGetMonikerFromCookie(cookie);
                if (string.IsNullOrWhiteSpace(moniker))
                {
                    continue;
                }

                // Only consider SQL query documents
                if (!moniker.EndsWith(".sql", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame == null)
                {
                    continue;
                }

                string caption = TryReadFrameCaption(frame);
                TryGetConnectionInfo(frame, out string server, out string database);
                list.Add(new OpenDocumentInfo
                {
                    Cookie = cookie,
                    Frame = frame,
                    Caption = caption,
                    Moniker = moniker,
                    Server = server,
                    Database = database
                });
            }

            return list;
        }

        private string TryGetMonikerFromCookie(uint docCookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (docCookie == 0) return null;

            try
            {
                rdt.GetDocumentInfo(
                    docCookie,
                    out uint _,
                    out uint _,
                    out uint _,
                    out string moniker,
                    out IVsHierarchy _,
                    out uint _,
                    out IntPtr _);

                return moniker;
            }
            catch
            {
                return null;
            }
        }

        private static string TryReadFrameCaption(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null) return null;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_Caption, out object caption) == VSConstants.S_OK)
                {
                    return caption as string;
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private static string TryReadFrameEditorCaption(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null) return null;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_EditorCaption, out object caption) == VSConstants.S_OK)
                {
                    return caption as string;
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private bool TryGetMonikerFromFrame(IVsWindowFrame frame, out string moniker)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            moniker = null;
            if (frame == null) return false;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out object mk) == VSConstants.S_OK)
                {
                    moniker = mk as string;
                    return !string.IsNullOrWhiteSpace(moniker);
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        private IVsWindowFrame TryGetFrameFromMoniker(string moniker)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (shellOpenDoc == null || string.IsNullOrWhiteSpace(moniker)) return null;

            // Main attempt: IsDocumentOpen
            try
            {
                Guid logicalView = Guid.Empty;
                uint[] itemid = new uint[1];
                if (shellOpenDoc.IsDocumentOpen(null, 0, moniker, ref logicalView, 0, out IVsUIHierarchy _, itemid, out IVsWindowFrame frame, out int isOpen) == VSConstants.S_OK && isOpen != 0)
                {
                    return frame;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Verbose($"TryGetFrameFromMoniker - Error: {ex.Message}");
                return null;
            }

            return null;
        }

        private static bool IsRenameEligible(string moniker, string caption)
        {
            if (string.IsNullOrWhiteSpace(moniker)) return false;

            // Check if it's a Temp file (New Query)
            if (IsTempFile(moniker))
            {
                // Verify it's a SQL file
                return moniker.EndsWith(".sql", StringComparison.OrdinalIgnoreCase);
            }

            // For saved files, we are eligible ONLY if the current caption matches the filename exactly.
            // This prevents overwriting our own renames or user-customized renames.
            try
            {
                string fileName = System.IO.Path.GetFileName(moniker);
                if (string.Equals(caption, fileName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Verbose($"IsRenameEligible - File name parse failed: {ex.Message}");
            }

            return false;
        }

        public static bool IsTempFile(string path)
        {
            try
            {
                string tempPath = System.IO.Path.GetTempPath();
                if (path != null && path.StartsWith(tempPath, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Verbose($"IsTempFile - Temp path check failed: {ex.Message}");
            }
            return false;
        }

        private static bool IsSqlDocumentMoniker(string moniker)
        {
            if (string.IsNullOrWhiteSpace(moniker))
            {
                return false;
            }

            if (!Path.IsPathRooted(moniker))
            {
                return false;
            }

            return moniker.EndsWith(".sql", StringComparison.OrdinalIgnoreCase);
        }
    }
}
