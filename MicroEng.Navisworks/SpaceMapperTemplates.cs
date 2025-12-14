using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;

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

                using var stream = File.OpenRead(_filePath);
                var serializer = new DataContractJsonSerializer(typeof(List<SpaceMapperTemplate>));
                return serializer.ReadObject(stream) as List<SpaceMapperTemplate> ?? new List<SpaceMapperTemplate> { CreateDefault() };
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
                        Name = "Level 0",
                        TargetType = SpaceMapperTargetType.SelectionTreeLevel,
                        MinTreeLevel = 0,
                        MaxTreeLevel = 0,
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
                ProcessingSettings = new SpaceMapperProcessingSettings()
            };
        }
    }
}
