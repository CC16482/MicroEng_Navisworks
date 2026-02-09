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

    public enum SpaceMapperTargetDefinition
    {
        EntireModel = 0,
        CurrentSelection = 1,
        SelectionSet = 2,
        SearchSet = 3,
        SelectionTreeLevel = 4
    }

    public sealed class SpaceMapperTargetDefinitionOption
    {
        public SpaceMapperTargetDefinition Value { get; set; }
        public string DisplayName { get; set; }

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(DisplayName) ? Value.ToString() : DisplayName;
        }
    }

    public enum SpaceMapperPerformancePreset
    {
        Fast = 0,
        Normal = 1,
        Accurate = 2,
        Auto = 3
    }

    internal enum SpaceMapperZoneBoundsMode
    {
        Aabb = 0,
        Obb = 1,
        KDop = 2,
        Hull = 3
    }

    internal enum SpaceMapperTargetBoundsMode
    {
        Midpoint = 0,
        Aabb = 1,
        Obb = 2,
        KDop = 3,
        Hull = 4
    }

    internal enum SpaceMapperMidpointMode
    {
        BoundingBoxCenter = 0,
        BoundingBoxBottomCenter = 1
    }

    internal enum SpaceMapperKDopVariant
    {
        KDop8 = 8,
        KDop14 = 14,
        KDop18 = 18
    }

    internal enum SpaceMapperZoneContainmentEngine
    {
        BoundsFast = 0,
        MeshAccurate = 1
    }

    internal enum SpaceMapperZoneResolutionStrategy
    {
        MostSpecific = 0,
        LargestOverlap = 1,
        FirstMatch = 2
    }

    internal enum SpaceMapperContainmentCalculationMode
    {
        Auto = 0,
        SamplePoints = 1,
        SamplePointsDense = 2,
        TargetGeometry = 3,
        BoundsOverlap = 4,
        TargetGeometryGpu = 5
    }

    public enum SpaceMapperFastTraversalMode
    {
        Auto = 0,
        ZoneMajor = 1,
        TargetMajor = 2
    }

    public enum SpaceMapperWritebackStrategy
    {
        VirtualNoBake = 0,
        OptimizedSingleCategory = 1,
        LegacyPerMapping = 2
    }

    internal enum SpaceMapperBenchmarkMode
    {
        ComputeOnly = 0,
        SimulateWriteback = 1,
        FullWriteback = 2
    }

    public enum SpaceMembershipMode
    {
        ContainedOnly,
        PartialOnly,
        ContainedAndPartial
    }

    public enum MultiZoneCombineMode
    {
        First = 0,
        Concatenate = 1,
        Min = 2,
        Max = 3,
        Average = 4,
        // Writes separate properties: Property, Property(1), Property(2)...
        Sequence = 5
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

    [DataContract]
    internal class SpaceMapperProcessingSettings
    {
        [DataMember(Order = 0)]
        public SpaceMapperProcessingMode ProcessingMode { get; set; } = SpaceMapperProcessingMode.Auto;

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

        [DataMember(Order = 12)]
        public int IndexGranularity { get; set; } = 0;

        [DataMember(Order = 13)]
        public SpaceMapperPerformancePreset PerformancePreset { get; set; } = SpaceMapperPerformancePreset.Auto;

        [DataMember(Order = 14)]
        public string ZoneBehaviorCategory { get; set; } = "ME_SpaceInfo";

        [DataMember(Order = 15)]
        public string ZoneBehaviorPropertyName { get; set; } = "Zone Behaviour";

        [DataMember(Order = 16)]
        public string ZoneBehaviorContainedValue { get; set; } = "Contained";

        [DataMember(Order = 17)]
        public string ZoneBehaviorPartialValue { get; set; } = "Partial";

        [DataMember(Order = 18)]
        public bool UseOriginPointOnly { get; set; }

        [DataMember(Order = 19)]
        public SpaceMapperFastTraversalMode FastTraversalMode { get; set; } = SpaceMapperFastTraversalMode.Auto;

        [DataMember(Order = 20)]
        public SpaceMapperWritebackStrategy WritebackStrategy { get; set; } = SpaceMapperWritebackStrategy.OptimizedSingleCategory;

        [DataMember(Order = 21)]
        public bool ShowInternalPropertiesDuringWriteback { get; set; }

        [DataMember(Order = 22)]
        public bool CloseDockPanesDuringRun { get; set; }

        [DataMember(Order = 23)]
        public bool SkipUnchangedWriteback { get; set; }

        [DataMember(Order = 24)]
        public bool PackWritebackProperties { get; set; }

        [DataMember(Order = 25)]
        public SpaceMapperZoneBoundsMode ZoneBoundsMode { get; set; } = SpaceMapperZoneBoundsMode.Aabb;

        [DataMember(Order = 26)]
        public SpaceMapperKDopVariant ZoneKDopVariant { get; set; } = SpaceMapperKDopVariant.KDop14;

        [DataMember(Order = 27)]
        public SpaceMapperTargetBoundsMode TargetBoundsMode { get; set; } = SpaceMapperTargetBoundsMode.Aabb;

        [DataMember(Order = 28)]
        public SpaceMapperKDopVariant TargetKDopVariant { get; set; } = SpaceMapperKDopVariant.KDop14;

        [DataMember(Order = 29)]
        public SpaceMapperMidpointMode TargetMidpointMode { get; set; } = SpaceMapperMidpointMode.BoundingBoxCenter;

        [DataMember(Order = 30)]
        public SpaceMapperZoneContainmentEngine ZoneContainmentEngine { get; set; } = SpaceMapperZoneContainmentEngine.BoundsFast;

        [DataMember(Order = 31)]
        public SpaceMapperZoneResolutionStrategy ZoneResolutionStrategy { get; set; } = SpaceMapperZoneResolutionStrategy.MostSpecific;

        [DataMember(Order = 32)]
        public bool ExcludeZonesFromTargets { get; set; }

        [DataMember(Order = 33)]
        public bool WriteZoneBehaviorProperty { get; set; }

        [DataMember(Order = 34)]
        public bool WriteZoneContainmentPercentProperty { get; set; }

        [DataMember(Order = 35)]
        public SpaceMapperContainmentCalculationMode ContainmentCalculationMode { get; set; } = SpaceMapperContainmentCalculationMode.Auto;

        [DataMember(Order = 36)]
        public int GpuRayCount { get; set; } = 2;

        [DataMember(Order = 37)]
        public double DockPaneCloseDelaySeconds { get; set; } = 2;

        [DataMember(Order = 38)]
        public bool EnableZoneOffsets { get; set; } = false;

        [DataMember(Order = 39)]
        public bool EnableOffsetAreaPass { get; set; } = false;

        [DataMember(Order = 40)]
        public bool WriteZoneOffsetMatchProperty { get; set; } = false;
    }

    [DataContract]
    public class SpaceMapperTargetRule
    {
        [DataMember(Order = 0)]
        public string Name { get; set; } = "Rule";

        [DataMember(Order = 1)]
        public SpaceMapperTargetDefinition TargetDefinition { get; set; } = SpaceMapperTargetDefinition.EntireModel;

        [DataMember(Order = 2)]
        public int? MinLevel { get; set; }

        [DataMember(Order = 3)]
        public int? MaxLevel { get; set; }

        [DataMember(Order = 4)]
        public string SetSearchName { get; set; }

        [DataMember(Order = 5)]
        public string CategoryFilter { get; set; }

        [DataMember(Order = 6)]
        public SpaceMembershipMode MembershipMode { get; set; } = SpaceMembershipMode.ContainedAndPartial;

        [DataMember(Order = 7)]
        public bool Enabled { get; set; } = true;
    }

    [DataContract]
    public class SpaceMapperMappingDefinition
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
        public Autodesk.Navisworks.Api.BoundingBox3D RawBoundingBox { get; set; }
        public Autodesk.Navisworks.Api.BoundingBox3D BoundingBox { get; set; }
        public List<Autodesk.Navisworks.Api.Vector3D> Vertices { get; set; } = new();
        public List<PlaneEquation> Planes { get; set; } = new();
        public IReadOnlyList<Autodesk.Navisworks.Api.Vector3D> TriangleVertices { get; set; }
        public int TriangleCount { get; set; }
        public bool HasTriangleMesh { get; set; }
        public bool MeshExtractionFailed { get; set; }
        public string MeshExtractionError { get; set; }
        public bool MeshIsClosed { get; set; }
        public int MeshBoundaryEdgeCount { get; set; }
        public int MeshNonManifoldEdgeCount { get; set; }
        public string MeshFallbackReason { get; set; }
        public string MeshFallbackDetail { get; set; }
    }

    internal class TargetGeometry
    {
        public string ItemKey { get; set; }
        public string DisplayName { get; set; }
        public Autodesk.Navisworks.Api.ModelItem ModelItem { get; set; }
        public Autodesk.Navisworks.Api.BoundingBox3D BoundingBox { get; set; }
        public List<Autodesk.Navisworks.Api.Vector3D> Vertices { get; set; } = new();
        public IReadOnlyList<Autodesk.Navisworks.Api.Vector3D> TriangleVertices { get; set; }
        public int TriangleCount { get; set; }
    }

    internal class SpaceMapperResolvedItem
    {
        public string ItemKey { get; set; }
        public string DisplayName { get; set; }
        public Autodesk.Navisworks.Api.ModelItem ModelItem { get; set; }
    }

    internal class ZoneTargetIntersection
    {
        public string ZoneId { get; set; }
        public string TargetItemKey { get; set; }
        public bool IsContained { get; set; }
        public bool IsPartial { get; set; }
        public double OverlapVolume { get; set; }
        public bool IsOffsetOnly { get; set; }
        [DataMember(Order = 5, EmitDefaultValue = false)]
        public double? ContainmentFraction { get; set; }
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

    internal sealed class SpaceMapperSlowZoneInfo
    {
        public int ZoneIndex { get; set; }
        public string ZoneId { get; set; }
        public string ZoneName { get; set; }
        public int CandidateCount { get; set; }
        public TimeSpan Elapsed { get; set; }
    }

    internal sealed class SpaceMapperGpuZoneDiagnostic
    {
        public int ZoneIndex { get; set; }
        public string ZoneId { get; set; }
        public string ZoneName { get; set; }
        public int CandidateCount { get; set; }
        public int EstimatedPoints { get; set; }
        public int TriangleCount { get; set; }
        public long WorkEstimate { get; set; }
        public int PointThreshold { get; set; }
        public long WorkThreshold { get; set; }
        public bool HasMesh { get; set; }
        public bool IsOpenMesh { get; set; }
        public bool AllowOpenMeshGpu { get; set; }
        public bool EligibleForGpu { get; set; }
        public bool UsedGpu { get; set; }
        public bool UsedOpenMeshRetry { get; set; }
        public bool PackedThresholds { get; set; }
        public string SkipReason { get; set; }
    }

    internal sealed class SpaceMapperEngineDiagnostics
    {
        public bool UsedPreflightIndex { get; set; }
        public SpaceMapperPerformancePreset PresetUsed { get; set; } = SpaceMapperPerformancePreset.Normal;
        public string TraversalUsed { get; set; }
        public long CandidatePairs { get; set; }
        public double AvgCandidatesPerZone { get; set; }
        public int MaxCandidatesPerZone { get; set; }
        public double AvgCandidatesPerTarget { get; set; }
        public int MaxCandidatesPerTarget { get; set; }
        public bool CandidateTargetStatsAvailable { get; set; }
        public int TargetsTotal { get; set; }
        public int TargetsWithBounds { get; set; }
        public int TargetsWithoutBounds { get; set; }
        public int TargetsSampled { get; set; }
        public int TargetsSampleSkippedNoBounds { get; set; }
        public int TargetsSampleSkippedNoGeometry { get; set; }
        public TimeSpan BuildIndexTime { get; set; }
        public TimeSpan CandidateQueryTime { get; set; }
        public TimeSpan NarrowPhaseTime { get; set; }
        public long MeshPointTests;
        public long BoundsPointTests;
        public long MeshFallbackPointTests;
        public string GpuBackend { get; set; }
        public string GpuInitFailureReason { get; set; }
        public int GpuZonesProcessed { get; set; }
        public long GpuPointsTested { get; set; }
        public long GpuTrianglesTested { get; set; }
        public TimeSpan GpuDispatchTime { get; set; }
        public TimeSpan GpuReadbackTime { get; set; }
        public string GpuAdapterName { get; set; }
        public string GpuAdapterLuid { get; set; }
        public int? GpuVendorId { get; set; }
        public int? GpuDeviceId { get; set; }
        public int? GpuSubSysId { get; set; }
        public int? GpuRevision { get; set; }
        public long? GpuDedicatedVideoMemory { get; set; }
        public long? GpuDedicatedSystemMemory { get; set; }
        public long? GpuSharedSystemMemory { get; set; }
        public string GpuFeatureLevel { get; set; }
        public int GpuPointThreshold { get; set; }
        public int GpuSamplePointsPerTarget { get; set; }
        public int GpuZonesEligible { get; set; }
        public int GpuZonesSkippedNoMesh { get; set; }
        public int GpuZonesSkippedMissingTriangles { get; set; }
        public int GpuZonesSkippedOpenMesh { get; set; }
        public int GpuZonesSkippedLowPoints { get; set; }
        public long GpuUncertainPoints { get; set; }
        public int GpuMaxTrianglesPerZone { get; set; }
        public int GpuMaxPointsPerZone { get; set; }
        public int GpuOpenMeshZonesEligible { get; set; }
        public int GpuOpenMeshZonesProcessed { get; set; }
        public int GpuOpenMeshBoundaryEdgeLimit { get; set; }
        public int GpuOpenMeshNonManifoldEdgeLimit { get; set; }
        public int GpuOpenMeshOutsideTolerance { get; set; }
        public double GpuOpenMeshNudge { get; set; }
        public int GpuBatchDispatchCount { get; set; }
        public int GpuBatchMaxZones { get; set; }
        public int GpuBatchMaxPoints { get; set; }
        public int GpuBatchMaxTriangles { get; set; }
        public double GpuBatchAvgZonesPerDispatch { get; set; }
        public List<SpaceMapperGpuZoneDiagnostic> GpuZoneDiagnostics { get; set; } = new();
        public double SlowZoneThresholdSeconds { get; set; }
        public List<SpaceMapperSlowZoneInfo> SlowZones { get; set; } = new();
    }

    internal class SpaceMapperRunStats
    {
        public int ZonesProcessed { get; set; }
        public int TargetsProcessed { get; set; }
        public int ZonesWithMesh { get; set; }
        public int ZonesMeshFallback { get; set; }
        public int MeshExtractionErrors { get; set; }
        public long MeshPointTests { get; set; }
        public long BoundsPointTests { get; set; }
        public long MeshFallbackPointTests { get; set; }
        public string GpuBackend { get; set; }
        public string GpuInitFailureReason { get; set; }
        public int GpuZonesProcessed { get; set; }
        public long GpuPointsTested { get; set; }
        public long GpuTrianglesTested { get; set; }
        public TimeSpan GpuDispatchTime { get; set; }
        public TimeSpan GpuReadbackTime { get; set; }
        public string GpuAdapterName { get; set; }
        public string GpuAdapterLuid { get; set; }
        public int? GpuVendorId { get; set; }
        public int? GpuDeviceId { get; set; }
        public int? GpuSubSysId { get; set; }
        public int? GpuRevision { get; set; }
        public long? GpuDedicatedVideoMemory { get; set; }
        public long? GpuDedicatedSystemMemory { get; set; }
        public long? GpuSharedSystemMemory { get; set; }
        public string GpuFeatureLevel { get; set; }
        public int GpuPointThreshold { get; set; }
        public int GpuSamplePointsPerTarget { get; set; }
        public int GpuZonesEligible { get; set; }
        public int GpuZonesSkippedNoMesh { get; set; }
        public int GpuZonesSkippedMissingTriangles { get; set; }
        public int GpuZonesSkippedOpenMesh { get; set; }
        public int GpuZonesSkippedLowPoints { get; set; }
        public long GpuUncertainPoints { get; set; }
        public int GpuMaxTrianglesPerZone { get; set; }
        public int GpuMaxPointsPerZone { get; set; }
        public int GpuOpenMeshZonesEligible { get; set; }
        public int GpuOpenMeshZonesProcessed { get; set; }
        public int GpuOpenMeshBoundaryEdgeLimit { get; set; }
        public int GpuOpenMeshNonManifoldEdgeLimit { get; set; }
        public int GpuOpenMeshOutsideTolerance { get; set; }
        public double GpuOpenMeshNudge { get; set; }
        public int GpuBatchDispatchCount { get; set; }
        public int GpuBatchMaxZones { get; set; }
        public int GpuBatchMaxPoints { get; set; }
        public int GpuBatchMaxTriangles { get; set; }
        public double GpuBatchAvgZonesPerDispatch { get; set; }
        public List<SpaceMapperGpuZoneDiagnostic> GpuZoneDiagnostics { get; set; } = new();
        public double SlowZoneThresholdSeconds { get; set; }
        public List<SpaceMapperSlowZoneInfo> SlowZones { get; set; } = new();
        public int ContainedTagged { get; set; }
        public int PartialTagged { get; set; }
        public int MultiZoneTagged { get; set; }
        public int Skipped { get; set; }
        public int SkippedUnchanged { get; set; }
        public string ModeUsed { get; set; }
        public TimeSpan Elapsed { get; set; }
        public SpaceMapperPerformancePreset PresetUsed { get; set; } = SpaceMapperPerformancePreset.Normal;
        public string TraversalUsed { get; set; }
        public long CandidatePairs { get; set; }
        public double AvgCandidatesPerZone { get; set; }
        public int MaxCandidatesPerZone { get; set; }
        public double AvgCandidatesPerTarget { get; set; }
        public int MaxCandidatesPerTarget { get; set; }
        public bool CandidateTargetStatsAvailable { get; set; }
        public int TargetsTotal { get; set; }
        public int TargetsWithBounds { get; set; }
        public int TargetsWithoutBounds { get; set; }
        public int TargetsSampled { get; set; }
        public int TargetsSampleSkippedNoBounds { get; set; }
        public int TargetsSampleSkippedNoGeometry { get; set; }
        public long WritesPerformed { get; set; }
        public int WritebackTargetsWritten { get; set; }
        public int WritebackCategoriesWritten { get; set; }
        public int WritebackPropertiesWritten { get; set; }
        public double AvgMsPerCategoryWrite { get; set; }
        public double AvgMsPerTargetWrite { get; set; }
        public SpaceMapperWritebackStrategy WritebackStrategy { get; set; } = SpaceMapperWritebackStrategy.OptimizedSingleCategory;
        public bool UsedPreflightIndex { get; set; }
        public TimeSpan ResolveTime { get; set; }
        public TimeSpan BuildGeometryTime { get; set; }
        public TimeSpan BuildIndexTime { get; set; }
        public TimeSpan CandidateQueryTime { get; set; }
        public TimeSpan NarrowPhaseTime { get; set; }
        public TimeSpan WriteBackTime { get; set; }
        public List<ZoneSummary> ZoneSummaries { get; set; } = new();
    }

    public class ZoneSummary
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

        [DataMember(Order = 4)]
        public string PreferredScraperProfileName { get; set; } = "Default";

        [DataMember(Order = 5)]
        public ZoneSourceType ZoneSource { get; set; } = ZoneSourceType.DataScraperZones;

        [DataMember(Order = 6)]
        public string ZoneSetName { get; set; }
    }
}
