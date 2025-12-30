using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace MicroEng.Navisworks
{
    internal class SpaceMapperTemplateStore
    {
        private readonly string _filePath;

        public SpaceMapperTemplateStore(string baseDir)
        {
            _filePath = Path.Combine(baseDir ?? AppDomain.CurrentDomain.BaseDirectory, "space_mapper_templates.json");
        }

        public List<SpaceMapperTemplate> Load()
        {
            try
            {
                if (!File.Exists(_filePath))
                {
                    return new List<SpaceMapperTemplate> { CreateDefault() };
                }

                var json = File.ReadAllText(_filePath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return new List<SpaceMapperTemplate> { CreateDefault() };
                }

                if (LooksLegacy(json))
                {
                    var legacyTemplates = Deserialize<List<LegacySpaceMapperTemplate>>(json);
                    var migrated = legacyTemplates?.Select(MigrateLegacyTemplate).ToList() ?? new List<SpaceMapperTemplate>();
                    return EnsureDefaults(migrated);
                }

                var templates = Deserialize<List<SpaceMapperTemplate>>(json);
                return EnsureDefaults(templates);
            }
            catch
            {
                return new List<SpaceMapperTemplate> { CreateDefault() };
            }
        }

        public void Save(IEnumerable<SpaceMapperTemplate> templates)
        {
            try
            {
                using var stream = File.Create(_filePath);
                var serializer = new DataContractJsonSerializer(typeof(List<SpaceMapperTemplate>));
                serializer.WriteObject(stream, templates?.ToList() ?? new List<SpaceMapperTemplate>());
            }
            catch
            {
                // non-fatal
            }
        }

        public static SpaceMapperTemplate CreateDefault()
        {
            return new SpaceMapperTemplate
            {
                Name = "Default",
                TargetRules = new List<SpaceMapperTargetRule>
                {
                    new SpaceMapperTargetRule
                    {
                        Name = "All Targets",
                        TargetDefinition = SpaceMapperTargetDefinition.EntireModel,
                        MembershipMode = SpaceMembershipMode.ContainedAndPartial
                    }
                },
                Mappings = new List<SpaceMapperMappingDefinition>
                {
                    new SpaceMapperMappingDefinition
                    {
                        Name = "Zone Name",
                        TargetPropertyName = "Zone Name",
                        ZoneCategory = "Zone",
                        ZonePropertyName = "Name"
                    }
                },
                ProcessingSettings = new SpaceMapperProcessingSettings(),
                PreferredScraperProfileName = "Default",
                ZoneSource = ZoneSourceType.DataScraperZones
            };
        }

        private static bool LooksLegacy(string json)
        {
            return json.IndexOf("\"TargetType\"", StringComparison.OrdinalIgnoreCase) >= 0
                || json.IndexOf("\"TargetSource\"", StringComparison.OrdinalIgnoreCase) >= 0
                || json.IndexOf("\"TargetSetName\"", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static T Deserialize<T>(string json)
        {
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(json));
            var serializer = new DataContractJsonSerializer(typeof(T));
            return (T)serializer.ReadObject(stream);
        }

        private static List<SpaceMapperTemplate> EnsureDefaults(List<SpaceMapperTemplate> templates)
        {
            if (templates == null || templates.Count == 0)
            {
                return new List<SpaceMapperTemplate> { CreateDefault() };
            }

            foreach (var template in templates)
            {
                if (template.TargetRules == null || template.TargetRules.Count == 0)
                {
                    template.TargetRules = new List<SpaceMapperTargetRule>
                    {
                        new SpaceMapperTargetRule
                        {
                            Name = "All Targets",
                            TargetDefinition = SpaceMapperTargetDefinition.EntireModel,
                            MembershipMode = SpaceMembershipMode.ContainedAndPartial,
                            Enabled = true
                        }
                    };
                }
            }

            return templates;
        }

        private static SpaceMapperTemplate MigrateLegacyTemplate(LegacySpaceMapperTemplate legacy)
        {
            var template = new SpaceMapperTemplate
            {
                Name = legacy?.Name ?? "Default",
                Mappings = legacy?.Mappings ?? new List<SpaceMapperMappingDefinition>(),
                ProcessingSettings = legacy?.ProcessingSettings ?? new SpaceMapperProcessingSettings(),
                PreferredScraperProfileName = legacy?.PreferredScraperProfileName ?? "Default",
                ZoneSource = legacy?.ZoneSource ?? ZoneSourceType.DataScraperZones,
                ZoneSetName = legacy?.ZoneSetName
            };

            var rules = new List<SpaceMapperTargetRule>();
            if (legacy?.TargetRules != null && legacy.TargetRules.Count > 0)
            {
                foreach (var rule in legacy.TargetRules)
                {
                    var mapped = new SpaceMapperTargetRule
                    {
                        Name = rule.Name,
                        TargetDefinition = MapTargetDefinition(rule.TargetType),
                        MinLevel = rule.MinTreeLevel,
                        MaxLevel = rule.MaxTreeLevel,
                        SetSearchName = rule.SetName,
                        CategoryFilter = rule.CategoryFilter,
                        MembershipMode = rule.MembershipMode,
                        Enabled = rule.Enabled
                    };
                    NormalizeRule(mapped);
                    rules.Add(mapped);
                }
            }

            if (rules.Count == 0 && legacy != null)
            {
                var fromSource = new SpaceMapperTargetRule
                {
                    Name = "All Targets",
                    TargetDefinition = MapTargetSource(legacy.TargetSource),
                    SetSearchName = legacy.TargetSetName,
                    MembershipMode = SpaceMembershipMode.ContainedAndPartial,
                    Enabled = true
                };
                NormalizeRule(fromSource);
                rules.Add(fromSource);
            }

            template.TargetRules = rules;
            return template;
        }

        private static SpaceMapperTargetDefinition MapTargetDefinition(LegacySpaceMapperTargetType type)
        {
            return type switch
            {
                LegacySpaceMapperTargetType.SelectionTreeLevel => SpaceMapperTargetDefinition.SelectionTreeLevel,
                LegacySpaceMapperTargetType.SelectionSet => SpaceMapperTargetDefinition.SelectionSet,
                LegacySpaceMapperTargetType.SearchSet => SpaceMapperTargetDefinition.SelectionSet,
                LegacySpaceMapperTargetType.CurrentSelection => SpaceMapperTargetDefinition.CurrentSelection,
                _ => SpaceMapperTargetDefinition.EntireModel
            };
        }

        private static SpaceMapperTargetDefinition MapTargetSource(LegacyTargetSourceType source)
        {
            return source switch
            {
                LegacyTargetSourceType.SelectionSet => SpaceMapperTargetDefinition.SelectionSet,
                LegacyTargetSourceType.SearchSet => SpaceMapperTargetDefinition.SelectionSet,
                LegacyTargetSourceType.EntireModel => SpaceMapperTargetDefinition.EntireModel,
                _ => SpaceMapperTargetDefinition.EntireModel
            };
        }

        private static void NormalizeRule(SpaceMapperTargetRule rule)
        {
            if (rule == null)
            {
                return;
            }

            if (rule.TargetDefinition == SpaceMapperTargetDefinition.SearchSet)
            {
                rule.TargetDefinition = SpaceMapperTargetDefinition.SelectionSet;
            }

            switch (rule.TargetDefinition)
            {
                case SpaceMapperTargetDefinition.SelectionSet:
                    rule.MinLevel = null;
                    rule.MaxLevel = null;
                    break;
                case SpaceMapperTargetDefinition.SelectionTreeLevel:
                    if (rule.MinLevel == null && rule.MaxLevel == null)
                    {
                        rule.MinLevel = 0;
                        rule.MaxLevel = 0;
                    }
                    else if (rule.MinLevel == null)
                    {
                        rule.MinLevel = rule.MaxLevel;
                    }
                    else if (rule.MaxLevel == null)
                    {
                        rule.MaxLevel = rule.MinLevel;
                    }
                    break;
                default:
                    rule.SetSearchName = null;
                    rule.MinLevel = null;
                    rule.MaxLevel = null;
                    break;
            }
        }
    }

    [DataContract]
    internal class LegacySpaceMapperTemplate
    {
        [DataMember(Order = 0)]
        public string Name { get; set; } = "Default";

        [DataMember(Order = 1)]
        public List<LegacySpaceMapperTargetRule> TargetRules { get; set; } = new();

        [DataMember(Order = 2)]
        public List<SpaceMapperMappingDefinition> Mappings { get; set; } = new();

        [DataMember(Order = 3)]
        public SpaceMapperProcessingSettings ProcessingSettings { get; set; } = new();

        [DataMember(Order = 4)]
        public string PreferredScraperProfileName { get; set; } = "Default";

        [DataMember(Order = 5)]
        public ZoneSourceType ZoneSource { get; set; } = ZoneSourceType.DataScraperZones;

        [DataMember(Order = 6)]
        public string ZoneSetName { get; set; }

        [DataMember(Order = 7)]
        public LegacyTargetSourceType TargetSource { get; set; } = LegacyTargetSourceType.EntireModel;

        [DataMember(Order = 8)]
        public string TargetSetName { get; set; }
    }

    [DataContract]
    internal class LegacySpaceMapperTargetRule
    {
        [DataMember(Order = 0)]
        public string Name { get; set; } = "Rule";

        [DataMember(Order = 1)]
        public LegacySpaceMapperTargetType TargetType { get; set; } = LegacySpaceMapperTargetType.SelectionTreeLevel;

        [DataMember(Order = 2)]
        public int? MinTreeLevel { get; set; } = 0;

        [DataMember(Order = 3)]
        public int? MaxTreeLevel { get; set; } = 0;

        [DataMember(Order = 4)]
        public string SetName { get; set; }

        [DataMember(Order = 5)]
        public string CategoryFilter { get; set; }

        [DataMember(Order = 6)]
        public SpaceMembershipMode MembershipMode { get; set; } = SpaceMembershipMode.ContainedAndPartial;

        [DataMember(Order = 7)]
        public bool Enabled { get; set; } = true;
    }

    internal enum LegacySpaceMapperTargetType
    {
        SelectionTreeLevel,
        SelectionSet,
        SearchSet,
        CurrentSelection,
        VisibleInView
    }

    internal enum LegacyTargetSourceType
    {
        EntireModel,
        Visible,
        Hidden,
        SelectionSet,
        SearchSet,
        ViewpointVisible
    }
}
