using System;
using System.Collections.Generic;
using System.Linq;

namespace MicroEng.Navisworks.SmartSets
{
    public sealed class SmartSetPackDefinition
    {
        public SmartSetPackDefinition(string name, string description, IEnumerable<SmartSetRule> rules)
        {
            Name = name ?? "";
            Description = description ?? "";
            Rules = rules?.ToList() ?? new List<SmartSetRule>();
        }

        public string Name { get; }
        public string Description { get; }
        public IReadOnlyList<SmartSetRule> Rules { get; }

        public IEnumerable<SmartSetRecipe> BuildRecipes(string profile)
        {
            var recipe = new SmartSetRecipe
            {
                Name = Name,
                Description = Description,
                DataScraperProfile = profile ?? "",
                OutputType = SmartSetOutputType.SearchSet,
                FolderPath = string.IsNullOrWhiteSpace(profile)
                    ? "MicroEng/Smart Sets/Packs"
                    : $"MicroEng/Smart Sets/{profile}/Packs"
            };

            foreach (var rule in Rules)
            {
                if (rule == null) continue;
                recipe.Rules.Add(new SmartSetRule
                {
                    GroupId = rule.GroupId,
                    Category = rule.Category,
                    Property = rule.Property,
                    Operator = rule.Operator,
                    Value = rule.Value,
                    Enabled = rule.Enabled
                });
            }

            return new[] { recipe };
        }

        public List<string> CheckMissingProperties(IDataScrapeSessionView session)
        {
            var missing = new List<string>();
            if (session?.Properties == null)
            {
                return missing;
            }

            var keys = new HashSet<string>(
                session.Properties.Select(p => $"{p.Category}::{p.Name}"),
                StringComparer.OrdinalIgnoreCase);

            foreach (var rule in Rules)
            {
                if (rule == null) continue;

                var key = $"{rule.Category}::{rule.Property}";
                if (!keys.Contains(key))
                {
                    missing.Add(key);
                }
            }

            return missing.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        public static SmartSetRule MakeRule(string category, string property, SmartSetOperator op, string value)
        {
            return new SmartSetRule
            {
                Category = category ?? "",
                Property = property ?? "",
                Operator = op,
                Value = value ?? "",
                Enabled = true
            };
        }
    }
}
