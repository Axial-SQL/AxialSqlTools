using System.Collections.Generic;
using System.Runtime.Serialization;

namespace AxialSqlTools.TabColoring
{
    [DataContract]
    internal sealed class TabColoringConfig
    {
        [DataMember(Name = "description", IsRequired = false, Order = 0)]
        public string Description { get; set; }

        [DataMember(Name = "settings", IsRequired = false, Order = 1)]
        public TabColoringSettings Settings { get; set; } = new TabColoringSettings();

        [DataMember(Name = "queryRules", IsRequired = false, Order = 2)]
        public List<TabColoringQueryRule> QueryRules { get; set; } = new List<TabColoringQueryRule>();
    }

    [DataContract]
    internal sealed class TabColoringSettings
    {
        [DataMember(Name = "enabled", IsRequired = false, Order = 0)]
        public bool Enabled { get; set; } = false;

        [DataMember(Name = "queryTimeoutSeconds", IsRequired = false, Order = 1)]
        public int QueryTimeoutSeconds { get; set; } = 5;

        [DataMember(Name = "defaultColorIndex", IsRequired = false, Order = 2)]
        public int? DefaultColorIndex { get; set; }
    }

    [DataContract]
    internal sealed class TabColoringQueryRule
    {
        [DataMember(Name = "name", IsRequired = false, Order = 0)]
        public string Name { get; set; }

        [DataMember(Name = "query", IsRequired = true, Order = 1)]
        public string Query { get; set; }

        [DataMember(Name = "expectedValue", IsRequired = false, Order = 2)]
        public string ExpectedValue { get; set; }

        [DataMember(Name = "colorIndex", IsRequired = true, Order = 3)]
        public int ColorIndex { get; set; }

        [DataMember(Name = "priority", IsRequired = false, Order = 4)]
        public int Priority { get; set; } = 0;

        [DataMember(Name = "isEnabled", IsRequired = false, Order = 5)]
        public bool IsEnabled { get; set; } = true;
    }
}
