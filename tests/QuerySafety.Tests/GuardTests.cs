using System;
using AxialSqlTools;
using AxialSqlTools.QuerySafety;
using EnvDTE;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using static Program;

internal static class GuardTests
{
    private class Connection
    {
        public Guid ClientConnectionId { get; set; } = Guid.NewGuid();
        public string State { get; set; } = "Open";
        public string DataSource { get; set; } = "server";
        public string Database { get; set; } = "db";
    }
    private class Editor
    {
        private readonly object m_connection;
        public Editor(object connection) { m_connection = connection; }
    }
    private static DTE Host(string sql = "DELETE dbo.T", Connection connection = null)
    {
        var host = new DTE { ActiveDocument = new Document(), ActiveWindow = new Window { Object = new Editor(connection ?? new Connection()) } };
        host.ActiveDocument.TextDocument.StartPoint.Text = sql;
        return host;
    }
    private static void Case(string name, Action run)
    {
        Test(name, () => {
            FatalActionGuard.ResetApprovals(); SettingsManager.Enabled = true;
            ServiceCache.ScriptFactory = new Factory();
            FatalActionWarningDialog.Prompts = 0; FatalActionWarningDialog.Remember = false;
            FatalActionWarningDialog.Decision = () => false;
            run();
        });
    }
    internal static void Run()
    {
        Case("disabled guard does not block execution", () => { SettingsManager.Enabled=false; Check(!FatalActionGuard.ShouldCancel(Host()) && FatalActionWarningDialog.Prompts==0); });
        Case("no active document leaves other commands alone", () => Check(!FatalActionGuard.ShouldCancel(new DTE())));
        Case("cancel prevents execution", () => Check(FatalActionGuard.ShouldCancel(Host())));
        Case("dialog close prevents execution", () => { FatalActionWarningDialog.Decision=()=>null; Check(FatalActionGuard.ShouldCancel(Host())); });
        Case("Run once prompts again on next execution", () => { FatalActionWarningDialog.Decision=()=>true; var h=Host(); Check(!FatalActionGuard.ShouldCancel(h) && !FatalActionGuard.ShouldCancel(h)); Check(FatalActionWarningDialog.Prompts==2); });
        Case("Allow repeats prompts only on first execution", () => { FatalActionWarningDialog.Decision=()=>true; FatalActionWarningDialog.Remember=true; var h=Host(); Check(!FatalActionGuard.ShouldCancel(h) && !FatalActionGuard.ShouldCancel(h) && !FatalActionGuard.ShouldCancel(h)); Check(FatalActionWarningDialog.Prompts==1); });
        Case("edited SQL requires new approval", () => { FatalActionWarningDialog.Decision=()=>true; FatalActionWarningDialog.Remember=true; var h=Host(); FatalActionGuard.ShouldCancel(h); h.ActiveDocument.TextDocument.StartPoint.Text="TRUNCATE TABLE dbo.S"; FatalActionWarningDialog.Decision=()=>false; Check(FatalActionGuard.ShouldCancel(h)); });
        Case("connection reconnect requires new approval", () => { FatalActionWarningDialog.Decision=()=>true; FatalActionWarningDialog.Remember=true; var c=new Connection(); var h=Host(connection:c); FatalActionGuard.ShouldCancel(h); c.ClientConnectionId=Guid.NewGuid(); FatalActionWarningDialog.Decision=()=>false; Check(FatalActionGuard.ShouldCancel(h)); });
        Case("toolbar database change requires new approval", () => { FatalActionWarningDialog.Decision=()=>true; FatalActionWarningDialog.Remember=true; var h=Host(); FatalActionGuard.ShouldCancel(h); ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo.AdvancedOptions["DATABASE"]="production"; FatalActionWarningDialog.Decision=()=>false; Check(FatalActionGuard.ShouldCancel(h)); });
        Case("live database change requires new approval", () => { FatalActionWarningDialog.Decision=()=>true; FatalActionWarningDialog.Remember=true; var c=new Connection(); var h=Host(connection:c); FatalActionGuard.ShouldCancel(h); c.Database="production"; FatalActionWarningDialog.Decision=()=>false; Check(FatalActionGuard.ShouldCancel(h)); });
        Case("selected SQL overrides unsafe unselected SQL", () => { var h=Host(); h.ActiveDocument.Selection=new TextSelection { IsEmpty=false, Text="SELECT 1" }; Check(!FatalActionGuard.ShouldCancel(h) && FatalActionWarningDialog.Prompts==0); });
        Case("selected unsafe SQL is checked without unselected WHERE", () => { var h=Host("DELETE dbo.T WHERE id=1"); h.ActiveDocument.Selection=new TextSelection { IsEmpty=false, Text="DELETE dbo.T" }; Check(FatalActionGuard.ShouldCancel(h)); });
        Case("selected whitespace does not execute entire document", () => { var h=Host(); h.ActiveDocument.Selection=new TextSelection { IsEmpty=false, Text="  " }; Check(!FatalActionGuard.ShouldCancel(h) && FatalActionWarningDialog.Prompts==0); });
        Case("temp-only query executes without warning", () => Check(!FatalActionGuard.ShouldCancel(Host("DELETE #t; UPDATE @t SET x=1; TRUNCATE TABLE ##t"))));
        Case("unknown connection supports Run once only", () => { var c=new Connection { ClientConnectionId=Guid.Empty }; var h=Host(connection:c); FatalActionWarningDialog.Decision=()=>true; FatalActionWarningDialog.Remember=true; Check(!FatalActionGuard.ShouldCancel(h) && !FatalActionGuard.ShouldCancel(h)); Check(!FatalActionWarningDialog.CanRemember && FatalActionWarningDialog.Prompts==2); });
        Case("parse failures cannot receive repeat approval", () => { var h=Host("DELETE FROM dbo.T; nonsense syntax here"); FatalActionWarningDialog.Decision=()=>true; FatalActionWarningDialog.Remember=true; Check(!FatalActionGuard.ShouldCancel(h) && !FatalActionGuard.ShouldCancel(h)); Check(!FatalActionWarningDialog.CanRemember && FatalActionWarningDialog.Prompts==2); });
        Case("closing document clears approval", () => { var h=Host(); FatalActionWarningDialog.Decision=()=>true; FatalActionWarningDialog.Remember=true; FatalActionGuard.ShouldCancel(h); FatalActionGuard.ForgetDocument(h.ActiveDocument); FatalActionWarningDialog.Decision=()=>false; Check(FatalActionGuard.ShouldCancel(h)); });
        Case("saving settings clears approvals", () => { var h=Host(); FatalActionWarningDialog.Decision=()=>true; FatalActionWarningDialog.Remember=true; FatalActionGuard.ShouldCancel(h); FatalActionGuard.ResetApprovals(); FatalActionWarningDialog.Decision=()=>false; Check(FatalActionGuard.ShouldCancel(h)); });
        Case("unavailable text offers cancellation", () => { var h=Host(); h.ActiveDocument.TextDocument=null; Check(FatalActionGuard.ShouldCancel(h) && !FatalActionWarningDialog.CanRemember); });
        Case("dialog failure cancels and releases dialog state", () => { var h=Host(); FatalActionWarningDialog.Decision=()=>throw new Exception("Cannot display"); Check(FatalActionGuard.ShouldCancel(h)); FatalActionWarningDialog.Decision=()=>false; Check(FatalActionGuard.ShouldCancel(h) && FatalActionWarningDialog.Prompts==2); });
        Case("reentrant execution is cancelled", () => { var h=Host(); FatalActionWarningDialog.Decision=()=>{ Check(FatalActionGuard.ShouldCancel(h)); return true; }; Check(!FatalActionGuard.ShouldCancel(h) && FatalActionWarningDialog.Prompts==1); });
    }
}
