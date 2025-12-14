using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace MicroEng.Navisworks
{
    public enum SpaceMapperProcessingMode
    {
        Auto,
        CpuNormal,
        GpuQuick,
        GpuIntensive,
        Debug
    }

    public enum SpaceMapperTargetType
    {
        SelectionTreeLevel,
        SelectionSet,
        SearchSet,
        CurrentSelection,
        VisibleInView
    }

    public enum SpaceMembershipMode
    {
        ContainedOnly,
        PartialOnly,
        ContainedAndPartial
    }

    public enum MultiZoneCombineMode
    {
        First,
        Concatenate,
        Min,
        Max,
        Average
    }

    public enum WriteMode
    {
        Overwrite,
        OnlyIfBlank,
        Append
    }

    public enum SpaceMapperScope
    {
        EntireModel,
        CurrentView,
        CurrentSelection,
        SelectionSet,
        SearchSet
    }

    public enum ZoneSourceType
    {
        DataScraperZones,
        ZoneSelectionSet,
        ZoneSearchSet
    }

    public enum TargetSourceType
    {
        EntireModel,
        Visible,
        Hidden,
        SelectionSet,
        SearchSet,
        ViewpointVisible
    }

    [DataContract]
    internal class SpaceMapperProcessingSettings
    {
        [DataMember(Order = 0)]
        public SpaceMapperProcessingMode ProcessingMode { get; set; } = SpaceMapperProcessingMode.CpuNormal;

        [DataMember(Order = 1)]
        public bool TreatPartialAsContained { get; set; }

        [DataMember(Order = 2)]
        public bool TagPartialSeparately { get; set; }

        [DataMember(Order = 3)]
        public bool EnableMultipleZones { get; set; } = true;

        [DataMember(Order = 4)]
        public double Offset3D { get; set; }

        [DataMember(Order = 5)]
        public double OffsetTop { get; set; }

        [DataMember(Order = 6)]
        public double OffsetBottom { get; set; }

        [DataMember(Order = 7)]
        public double OffsetSides { get; set; }

        [DataMember(Order = 8)]
        public string Units { get; set; } = "Millimeters";

        [DataMember(Order = 9)]
        public string OffsetMode { get; set; } = "From Zone Geometry";

        [DataMember(Order = 10)]
        public int? MaxThreads { get; set; }

        [DataMember(Order = 11)]
        public int? BatchSize { get; set; }
    }

    [DataContract]
    internal class SpaceMapperTargetRule
    {
        [DataMember(Order = 0)]
        public string Name { get; set; } = "Rule";

        [DataMember(Order = 1)]
        public SpaceMapperTargetType TargetType { get; set; } = SpaceMapperTargetType.SelectionTreeLevel;

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

    [DataContract]
    internal class SpaceMapperMappingDefinition
    {
        [DataMember(Order = 0)]
        public string Name { get; set; } = "Mapping";

        [DataMember(Order = 1)]
        public string ZoneCategory { get; set; }

        [DataMember(Order = 2)]
        public string ZonePropertyName { get; set; }

        [DataMember(Order = 3)]
        public string TargetCategory { get; set; } = "ME_SpaceInfo";

        [DataMember(Order = 4)]
        public string TargetPropertyName { get; set; }

        [DataMember(Order = 5)]
        public WriteMode WriteMode { get; set; } = WriteMode.Overwrite;

        [DataMember(Order = 6)]
        public string AppendSeparator { get; set; } = ", ";

        [DataMember(Order = 7)]
        public string PartialFlagValue { get; set; } = "Partial";

        [DataMember(Order = 8)]
        public MultiZoneCombineMode MultiZoneCombineMode { get; set; } = MultiZoneCombineMode.First;

        [DataMember(Order = 9)]
        public bool IsEditable { get; set; } = true;
    }

    internal class ZoneGeometry
    {
        public string ZoneId { get; set; }
        public string DisplayName { get; set; }
        public Autodesk.Navisworks.Api.ModelItem ModelItem { get; set; }
        public Autodesk.Navisworks.Api.BoundingBox3D BoundingBox { get; set; }
        public List<Autodesk.Navisworks.Api.Vector3D> Vertices { get; set; } = new();
        public List<PlaneEquation> Planes { get; set; } = new();
    }

    internal class TargetGeometry
    {
        public string ItemKey { get; set; }
        public string DisplayName { get; set; }
        public Autodesk.Navisworks.Api.ModelItem ModelItem { get; set; }
        public Autodesk.Navisworks.Api.BoundingBox3D BoundingBox { get; set; }
        public List<Autodesk.Navisworks.Api.Vector3D> Vertices { get; set; } = new();
    }

    internal class ZoneTargetIntersection
    {
        public string ZoneId { get; set; }
        public string TargetItemKey { get; set; }
        public bool IsContained { get; set; }
        public bool IsPartial { get; set; }
        public double OverlapVolume { get; set; }
    }

    internal struct PlaneEquation
    {
        public Autodesk.Navisworks.Api.Vector3D Normal;
        public double D;
    }

    internal class SpaceMapperProgress
    {
        public int ProcessedPairs { get; set; }
        public int TotalPairs { get; set; }
        public int ZonesProcessed { get; set; }
        public int TargetsProcessed { get; set; }
        public int Percentage => TotalPairs <= 0 ? 0 : (int)(ProcessedPairs * 100.0 / TotalPairs);
    }

    internal class SpaceMapperRunStats
    {
        public int ZonesProcessed { get; set; }
        public int TargetsProcessed { get; set; }
        public int ContainedTagged { get; set; }
        public int PartialTagged { get; set; }
        public int MultiZoneTagged { get; set; }
        public int Skipped { get; set; }
        public string ModeUsed { get; set; }
        public TimeSpan Elapsed { get; set; }
        public List<ZoneSummary> ZoneSummaries { get; set; } = new();
    }

    internal class ZoneSummary
    {
        public string ZoneId { get; set; }
        public string ZoneName { get; set; }
        public int ContainedCount { get; set; }
        public int PartialCount { get; set; }
    }

    [DataContract]
    internal class SpaceMapperTemplate
    {
        [DataMember(Order = 0)]
        public string Name { get; set; } = "Default";

        [DataMember(Order = 1)]
        public List<SpaceMapperTargetRule> TargetRules { get; set; } = new();

        [DataMember(Order = 2)]
        public List<SpaceMapperMappingDefinition> Mappings { get; set; } = new();

        [DataMember(Order = 3)]
        public SpaceMapperProcessingSettings ProcessingSettings { get; set; } = new();
    }
}
