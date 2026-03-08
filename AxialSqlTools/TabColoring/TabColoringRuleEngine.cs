using System;
using System.Collections.Generic;
using System.Linq;

namespace AxialSqlTools.TabColoring
{
    internal static class TabColoringRuleEngine
    {
        public static IReadOnlyList<TabColoringQueryRule> GetEnabledRulesInPriorityOrder(TabColoringConfig config)
        {
            if (config?.QueryRules == null)
            {
                return Array.Empty<TabColoringQueryRule>();
            }

            return config.QueryRules
                .Where(r => r != null && r.IsEnabled && !string.IsNullOrWhiteSpace(r.Query))
                .OrderByDescending(r => r.Priority)
                .ToList();
        }

        public static bool IsMatch(TabColoringQueryRule rule, object queryResult)
        {
            if (rule == null)
            {
                return false;
            }

            if (queryResult == null || queryResult == DBNull.Value)
            {
                return string.IsNullOrEmpty(rule.ExpectedValue);
            }

            if (string.IsNullOrEmpty(rule.ExpectedValue))
            {
                return true;
            }

            return string.Equals(Convert.ToString(queryResult), rule.ExpectedValue, StringComparison.OrdinalIgnoreCase);
        }
    }
}
