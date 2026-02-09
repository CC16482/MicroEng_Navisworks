using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace MicroEng.Navisworks
{
    internal enum AppendValueMode
    {
        StaticValue,
        FromProperty,
        Expression
    }

    internal enum AppendValueOption
    {
        None,
        ConvertToDecimal,
        ConvertToInteger,
        ConvertToDate,
        FormatAsText,
        SumGroupProperty,
        UseParentProperty,
        UseParentRevitProperty,
        ReadAllPropertiesFromTab,
        ParseExcelFormula,
        PerformCalculation,
        ConvertFromRevit
    }

    internal enum ApplyPropertyTarget
    {
        Items,
        Groups,
        ItemsAndGroups
    }

    [DataContract]
    internal class AppendIntegrateRow
    {
        [DataMember(Order = 0)]
        public string TargetPropertyName { get; set; } = string.Empty;

        [DataMember(Order = 1)]
        public AppendValueMode Mode { get; set; } = AppendValueMode.StaticValue;

        [DataMember(Order = 2)]
        public string SourcePropertyPath { get; set; } = string.Empty;

        [DataMember(Order = 3)]
        public string SourcePropertyLabel { get; set; } = string.Empty;

        [DataMember(Order = 4)]
        public string StaticOrExpressionValue { get; set; } = string.Empty;

        [DataMember(Order = 5)]
        public AppendValueOption Option { get; set; } = AppendValueOption.None;

        [DataMember(Order = 6)]
        public bool Enabled { get; set; } = true;
    }

    [DataContract]
    internal class AppendIntegrateTemplate
    {
        [DataMember(Order = 0)]
        public string Name { get; set; }

        [DataMember(Order = 1)]
        public string TargetTabName { get; set; } = "MicroEng";

        [DataMember(Order = 2)]
        public List<AppendIntegrateRow> Rows { get; set; } = new();

        [DataMember(Order = 3)]
        public ApplyPropertyTarget ApplyPropertyTo { get; set; } = ApplyPropertyTarget.Items;

        [DataMember(Order = 4)]
        public bool CreateTargetTabIfMissing { get; set; } = true;

        [DataMember(Order = 5)]
        public bool KeepExistingTabs { get; set; } = true;

        [DataMember(Order = 6)]
        public bool UpdateExistingTargetTab { get; set; } = true;

        [DataMember(Order = 7)]
        public bool DeletePropertyIfBlank { get; set; }

        [DataMember(Order = 8)]
        public bool DeleteTargetTabIfAllBlank { get; set; }

        [DataMember(Order = 9)]
        public bool ApplyToSelectionOnly { get; set; } = true;

        [DataMember(Order = 10)]
        public bool ShowInternalPropertyNames { get; set; }

        public static AppendIntegrateTemplate CreateDefault(string name)
        {
            return new AppendIntegrateTemplate
            {
                Name = name ?? "Default",
                TargetTabName = "MicroEng",
                ApplyPropertyTo = ApplyPropertyTarget.Items,
                CreateTargetTabIfMissing = true,
                KeepExistingTabs = true,
                UpdateExistingTargetTab = true,
                ApplyToSelectionOnly = true
            };
        }
    }

    internal class AppendIntegrateTemplateStore
    {
        private readonly string _templateFilePath;
        private readonly string _legacyTemplateFilePath;
        private bool _legacyMigrated;

        public AppendIntegrateTemplateStore(string baseDirectory)
        {
            _templateFilePath = MicroEngStorageSettings.GetDataFilePath("DataMapperTemplates.json");
            _legacyTemplateFilePath = Path.Combine(baseDirectory ?? AppDomain.CurrentDomain.BaseDirectory, "append_templates.json");
            Directory.CreateDirectory(Path.GetDirectoryName(_templateFilePath) ?? MicroEngStorageSettings.DataStorageDirectory);
        }

        public List<AppendIntegrateTemplate> Load()
        {
            try
            {
                var readPath = ResolveReadPath();
                if (!File.Exists(readPath))
                {
                    return new List<AppendIntegrateTemplate> { AppendIntegrateTemplate.CreateDefault("Default") };
                }

                using var stream = File.OpenRead(readPath);
                var serializer = new DataContractJsonSerializer(typeof(List<AppendIntegrateTemplate>));
                var result = serializer.ReadObject(stream) as List<AppendIntegrateTemplate> ??
                             new List<AppendIntegrateTemplate> { AppendIntegrateTemplate.CreateDefault("Default") };
                TryMigrateLegacy(readPath, result);
                return result;
            }
            catch
            {
                return new List<AppendIntegrateTemplate> { AppendIntegrateTemplate.CreateDefault("Default") };
            }
        }

        public void Save(IEnumerable<AppendIntegrateTemplate> templates)
        {
            var list = templates?.ToList() ?? new List<AppendIntegrateTemplate>();
            try
            {
                using var stream = File.Create(_templateFilePath);
                var serializer = new DataContractJsonSerializer(typeof(List<AppendIntegrateTemplate>));
                serializer.WriteObject(stream, list);
            }
            catch
            {
                // Persist failures are non-fatal; callers may surface a UI warning later.
            }
        }

        private string ResolveReadPath()
        {
            if (File.Exists(_templateFilePath))
            {
                return _templateFilePath;
            }

            if (File.Exists(_legacyTemplateFilePath))
            {
                return _legacyTemplateFilePath;
            }

            return _templateFilePath;
        }

        private void TryMigrateLegacy(string readPath, List<AppendIntegrateTemplate> templates)
        {
            if (_legacyMigrated)
            {
                return;
            }

            if (string.Equals(readPath, _legacyTemplateFilePath, StringComparison.OrdinalIgnoreCase))
            {
                Save(templates);
                _legacyMigrated = true;
            }
        }
    }
}
