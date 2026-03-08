using System.Collections.Generic;
using System.Runtime.Serialization;

namespace SSMS_EnvTabs
{
    [DataContract]
    internal sealed class TabGroupConfig
    {
        [DataMember(Name = "description", IsRequired = false, Order = 0)]
        public string Description { get; set; }

        [DataMember(Name = "help", IsRequired = false, Order = 1)]
        public string Help { get; set; }

        [DataMember(Name = "version", IsRequired = false, Order = 2)]
        public string Version { get; set; }

        [DataMember(Name = "settings", IsRequired = false, Order = 3)]
        public TabGroupSettings Settings { get; set; } = new TabGroupSettings();

        [DataMember(Name = "manualRegexLines", IsRequired = false, Order = 4)]
        public List<ManualRegexEntry> ManualRegexLines { get; set; } = new List<ManualRegexEntry>();

        [DataMember(Name = "serverAlias", IsRequired = false, Order = 5)]
        public Dictionary<string, string> ServerAliases { get; set; } = new Dictionary<string, string>();

        [DataMember(Name = "connectionGroups", IsRequired = false, Order = 6)]
        public List<TabGroupRule> ConnectionGroups { get; set; } = new List<TabGroupRule>();
    }

    [DataContract]
    internal sealed class ManualRegexEntry
    {
        [DataMember(Name = "groupName", IsRequired = false)]
        public string GroupName { get; set; }

        [DataMember(Name = "pattern", IsRequired = true)]
        public string Pattern { get; set; }

        [DataMember(Name = "priority", IsRequired = false)]
        public int Priority { get; set; } = 50;

        [DataMember(Name = "colorIndex", IsRequired = false)]
        public int? ColorIndex { get; set; }
    }


    [DataContract]
    internal sealed class TabGroupSettings
    {
        [DataMember(Name = "enableLogging", IsRequired = false, Order = 0)]
        public bool EnableLogging { get; set; } = true;

        [DataMember(Name = "enableVerboseLogging", IsRequired = false, Order = 0)]
        public bool EnableVerboseLogging { get; set; } = false;

        [DataMember(Name = "enableAutoRename", IsRequired = false, Order = 1)]
        public bool EnableAutoRename { get; set; } = true;

        [DataMember(Name = "enableAutoColor", IsRequired = false, Order = 2)]
        public bool EnableAutoColor { get; set; } = false;

        [DataMember(Name = "enableConfigurePrompt", IsRequired = false, Order = 3)]
        public bool EnableConfigurePrompt { get; set; } = true;

        [DataMember(Name = "enableConnectionPolling", IsRequired = false, Order = 4)]
        public bool EnableConnectionPolling { get; set; } = true;

        [DataMember(Name = "enableColorWarning", IsRequired = false, Order = 5)]
        public bool EnableColorWarning { get; set; } = true;

        [DataMember(Name = "enableServerAliasPrompt", IsRequired = false, Order = 6)]
        public bool EnableServerAliasPrompt { get; set; } = true;

        [DataMember(Name = "enableUpdateChecks", IsRequired = false, Order = 7)]
        public bool EnableUpdateChecks { get; set; } = true;

        [DataMember(Name = "autoConfigure", IsRequired = false, Order = 8)]
        public string AutoConfigure { get; set; }

        [DataMember(Name = "newQueryRenameStyle", IsRequired = false, Order = 9)]
        public string NewQueryRenameStyle { get; set; }

        [DataMember(Name = "suggestedGroupNameStyle", IsRequired = false, Order = 10)]
        public string SuggestedGroupNameStyle { get; set; }

        [DataMember(Name = "savedFileRenameStyle", IsRequired = false, Order = 11)]
        public string SavedFileRenameStyle { get; set; }

        [DataMember(Name = "enableRemoveDotSql", IsRequired = false, Order = 12)]
        public bool EnableRemoveDotSql { get; set; } = true;
    }

    [DataContract]
    internal sealed class TabGroupRule
    {
        [DataMember(Name = "groupName", IsRequired = false, Order = 0, EmitDefaultValue = true)]
        public string GroupName { get; set; }

        [DataMember(Name = "server", IsRequired = false, Order = 1)]
        public string Server { get; set; }

        [DataMember(Name = "database", IsRequired = false, Order = 2)]
        public string Database { get; set; }

        [DataMember(Name = "priority", IsRequired = false, Order = 3)]
        public int Priority { get; set; } = 0;

        [DataMember(Name = "colorIndex", IsRequired = false, Order = 4, EmitDefaultValue = true)]
        public int? ColorIndex { get; set; }
    }
}
