using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.CompilerServices;
using System.Text;
using Microsoft.Win32;
using Autodesk.Navisworks.Api;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace MicroEng.Navisworks
{
    internal static class SpaceMapperRunReportWriter
    {
        internal sealed class SpaceMapperLiveReportCache
        {
            public string ReportPath { get; set; }
            public string ModelTitle { get; set; }
            public string ModelPath { get; set; }
            public string GraphicsSnapshot { get; set; }
            public DateTimeOffset RunStartUtc { get; set; }
            public int Failures { get; set; }
            public bool Disabled { get; set; }
        }

        private static readonly Dictionary<string, string[]> TransformPropertyAliases = new(StringComparer.OrdinalIgnoreCase)
        {
            { "scalex", new[] { "scale0" } },
            { "scaley", new[] { "scale1" } },
            { "scalez", new[] { "scale2" } },
            { "rotationaxisx", new[] { "rotationaxis0", "rotaxis0" } },
            { "rotationaxisy", new[] { "rotationaxis1", "rotaxis1" } },
            { "rotationaxisz", new[] { "rotationaxis2", "rotaxis2" } },
            { "translationx", new[] { "translation0" } },
            { "translationy", new[] { "translation1" } },
            { "translationz", new[] { "translation2" } }
        };
        private const double PointMatchTolerance = 1e-3;

        internal static SpaceMapperLiveReportCache CreateLiveReportCache(SpaceMapperRequest request)
        {
            var doc = Application.ActiveDocument;
            var cache = new SpaceMapperLiveReportCache
            {
                ReportPath = BuildLiveReportPath(doc, request?.TemplateName),
                ModelTitle = GetDocumentTitle(doc),
                ModelPath = doc?.FileName ?? "<unsaved>",
                GraphicsSnapshot = BuildGraphicsSnapshot(),
                RunStartUtc = DateTimeOffset.UtcNow
            };
            return cache;
        }

        internal static void TryWriteLiveReport(
            SpaceMapperLiveReportCache cache,
            SpaceMapperRequest request,
            SpaceMapperResolvedData resolved,
            SpaceMapperComputeDataset dataset,
            SpaceMapperRunResult result,
            SpaceMapperRunProgressState progress)
        {
            if (cache == null || cache.Disabled || string.IsNullOrWhiteSpace(cache.ReportPath))
            {
                return;
            }

            try
            {
                var reportText = BuildLiveReport(cache, request, resolved, dataset, result, progress);
                Directory.CreateDirectory(Path.GetDirectoryName(cache.ReportPath));
                File.WriteAllText(cache.ReportPath, reportText, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                cache.Failures++;
                if (cache.Failures == 1)
                {
                    MicroEngActions.Log($"SpaceMapper live report write failed: {ex.Message}");
                }
                if (cache.Failures >= 3)
                {
                    cache.Disabled = true;
                }
            }
        }

        internal static string TryWriteReport(
            SpaceMapperRequest request,
            SpaceMapperResolvedData resolved,
            SpaceMapperComputeDataset dataset,
            SpaceMapperRunResult result)
        {
            return TryWriteReport(request, resolved, dataset, result, null);
        }

        internal static string TryWriteReport(
            SpaceMapperRequest request,
            SpaceMapperResolvedData resolved,
            SpaceMapperComputeDataset dataset,
            SpaceMapperRunResult result,
            SpaceMapperRunProgressState progress)
        {
            try
            {
                var reportText = BuildReport(request, resolved, dataset, result, progress);
                var reportPath = BuildReportPath(Application.ActiveDocument, request?.TemplateName);
                Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
                File.WriteAllText(reportPath, reportText, Encoding.UTF8);
                return reportPath;
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"SpaceMapper report write failed: {ex.Message}");
                return null;
            }
        }

        private static string BuildReport(
            SpaceMapperRequest request,
            SpaceMapperResolvedData resolved,
            SpaceMapperComputeDataset dataset,
            SpaceMapperRunResult result,
            SpaceMapperRunProgressState progress)
        {
            var stats = result?.Stats ?? new SpaceMapperRunStats();
            var settings = request?.ProcessingSettings ?? new SpaceMapperProcessingSettings();
            var doc = Application.ActiveDocument;
            var nowLocal = DateTime.Now;
            var nowUtc = DateTime.UtcNow;

            using var proc = Process.GetCurrentProcess();
            var sb = new StringBuilder();

            sb.AppendLine("# Space Mapper Run Report");
            sb.AppendLine($"Generated (UTC): {nowUtc:u}");
            sb.AppendLine($"Generated (Local): {nowLocal:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();

            sb.AppendLine("## Environment");
            sb.AppendLine($"- Model: {GetDocumentTitle(doc)}");
            sb.AppendLine($"- Model path: {doc?.FileName ?? "<unsaved>"}");
            sb.AppendLine($"- Navisworks API: {typeof(Application).Assembly.GetName().Version}");
            sb.AppendLine($"- Plugin version: {typeof(SpaceMapperService).Assembly.GetName().Version}");
            sb.AppendLine($"- OS: {Environment.OSVersion}");
            sb.AppendLine($"- Machine: {Environment.MachineName}");
            sb.AppendLine($"- CPU cores: {Environment.ProcessorCount}");
            sb.AppendLine($"- Working set (MB): {proc.WorkingSet64 / (1024.0 * 1024.0):0.0}");
            try
            {
                var startUtc = proc.StartTime.ToUniversalTime();
                sb.AppendLine($"- Process ID: {proc.Id}");
                sb.AppendLine($"- Process start (UTC): {startUtc:u}");
                sb.AppendLine($"- Process uptime (min): {(DateTime.UtcNow - startUtc).TotalMinutes:0.0}");
            }
            catch
            {
                sb.AppendLine("- Process info: <unavailable>");
            }
            sb.AppendLine($"- Is 64-bit process: {Environment.Is64BitProcess}");
            sb.AppendLine($"- CLR version: {Environment.Version}");
            sb.AppendLine($"- Base directory: {AppDomain.CurrentDomain.BaseDirectory}");
            sb.AppendLine();

            AppendGraphicsEnvironment(sb, stats);
            AppendProgressSnapshot(sb, progress);

            sb.AppendLine("## Step 1 - Setup");
            sb.AppendLine($"- Template: {request?.TemplateName ?? "<none>"}");
            sb.AppendLine($"- Scraper profile: {request?.ScraperProfileName ?? "<none>"}");
            sb.AppendLine($"- Scope: {request?.Scope}");
            sb.AppendLine($"- Zone source: {request?.ZoneSource}");
            sb.AppendLine($"- Zone set/search: {request?.ZoneSetName ?? "<none>"}");
            sb.AppendLine($"- Target rule count: {request?.TargetRules?.Count ?? 0}");
            sb.AppendLine($"- Mapping count: {request?.Mappings?.Count ?? 0}");
            sb.AppendLine();

            sb.AppendLine("## Step 2 - Zones & Targets");
            sb.AppendLine($"- Resolved zones (raw): {resolved?.ZoneModels?.Count ?? 0}");
            sb.AppendLine($"- Zones with bbox: {dataset?.Zones?.Count ?? 0}");
            sb.AppendLine($"- Resolved targets (raw): {resolved?.TargetModels?.Count ?? 0}");
            sb.AppendLine($"- Targets with bbox: {dataset?.TargetsForEngine?.Count ?? 0}");
            sb.AppendLine($"- Exclude zones from targets: {settings.ExcludeZonesFromTargets}");
            sb.AppendLine();

            sb.AppendLine("### Target Rules");
            var rules = request?.TargetRules ?? new List<SpaceMapperTargetRule>();
            var ruleCounts = BuildRuleCounts(resolved, rules);
            for (int i = 0; i < rules.Count; i++)
            {
                var rule = rules[i];
                var count = ruleCounts.TryGetValue(rule, out var hits) ? hits : 0;
                sb.AppendLine($"- {i + 1}. {rule.Name}");
                sb.AppendLine($"  - Definition: {rule.TargetDefinition}");
                sb.AppendLine($"  - Set/search: {rule.SetSearchName ?? "<none>"}");
                sb.AppendLine($"  - Category filter: {rule.CategoryFilter ?? "<none>"}");
                sb.AppendLine($"  - Levels: {FormatRange(rule.MinLevel, rule.MaxLevel)}");
                sb.AppendLine($"  - Membership: {rule.MembershipMode}");
                sb.AppendLine($"  - Enabled: {rule.Enabled}");
                sb.AppendLine($"  - Matched targets: {count}");
            }
            if (rules.Count == 0)
            {
                sb.AppendLine("- <none>");
            }
            sb.AppendLine();

            var intersectionTargets = BuildIntersectionTargetSet(result?.Intersections);
            AppendTargetTransformSection(sb, dataset, settings, intersectionTargets);

            sb.AppendLine("## Step 3 - Processing");
            sb.AppendLine($"- Processing mode: {settings.ProcessingMode}");
            sb.AppendLine($"- GPU ray count: {Math.Max(settings.GpuRayCount, 1)}");
            sb.AppendLine($"- Preset: {settings.PerformancePreset}");
            sb.AppendLine($"- Fast traversal: {settings.FastTraversalMode}");
            sb.AppendLine($"- Zone containment engine: {settings.ZoneContainmentEngine}");
            sb.AppendLine($"- Zone bounds: {settings.ZoneBoundsMode} (k-DOP: {settings.ZoneKDopVariant})");
            sb.AppendLine($"- Target bounds: {settings.TargetBoundsMode} (k-DOP: {settings.TargetKDopVariant}, midpoint: {settings.TargetMidpointMode})");
            sb.AppendLine($"- Use origin point only: {settings.UseOriginPointOnly}");
            sb.AppendLine($"- Resolution strategy: {settings.ZoneResolutionStrategy}");
            sb.AppendLine($"- Enable multiple zones: {settings.EnableMultipleZones}");
            sb.AppendLine($"- Treat partial as contained: {settings.TreatPartialAsContained}");
            sb.AppendLine($"- Tag partial separately: {settings.TagPartialSeparately}");
            sb.AppendLine($"- Write zone behavior property: {settings.WriteZoneBehaviorProperty}");
            sb.AppendLine($"- Write containment % property: {settings.WriteZoneContainmentPercentProperty}");
            sb.AppendLine($"- Containment calculation mode: {settings.ContainmentCalculationMode}");
            sb.AppendLine($"- Zone behavior category: {settings.ZoneBehaviorCategory}");
            sb.AppendLine($"- Zone behavior property: {settings.ZoneBehaviorPropertyName}");
            sb.AppendLine($"- Zone behavior contained value: {settings.ZoneBehaviorContainedValue}");
            sb.AppendLine($"- Zone behavior partial value: {settings.ZoneBehaviorPartialValue}");
            sb.AppendLine($"- Offset 3D: {settings.Offset3D}");
            sb.AppendLine($"- Offset sides: {settings.OffsetSides}");
            sb.AppendLine($"- Offset top: {settings.OffsetTop}");
            sb.AppendLine($"- Offset bottom: {settings.OffsetBottom}");
            sb.AppendLine($"- Units: {settings.Units}");
            sb.AppendLine($"- Offset mode: {settings.OffsetMode}");
            sb.AppendLine($"- Index granularity: {settings.IndexGranularity}");
            sb.AppendLine($"- Max threads: {settings.MaxThreads}");
            sb.AppendLine($"- Batch size: {settings.BatchSize}");
            sb.AppendLine($"- Writeback strategy: {settings.WritebackStrategy}");
            sb.AppendLine($"- Show internal properties during writeback: {settings.ShowInternalPropertiesDuringWriteback}");
            sb.AppendLine($"- Skip unchanged writeback: {settings.SkipUnchangedWriteback}");
            sb.AppendLine($"- Pack writeback outputs: {settings.PackWritebackProperties}");
            sb.AppendLine($"- Close dock panes during run: {settings.CloseDockPanesDuringRun}");
            sb.AppendLine($"- Zones with mesh: {stats.ZonesWithMesh}");
            var meshClosureWarnings = dataset?.Zones?
                .Where(z => z.HasTriangleMesh && !z.MeshIsClosed)
                .ToList() ?? new List<ZoneGeometry>();
            sb.AppendLine($"- Zones with non-watertight mesh (warning): {meshClosureWarnings.Count}");
            sb.AppendLine($"- Zones mesh fallback (mesh unusable): {stats.ZonesMeshFallback}");
            sb.AppendLine($"- Mesh extraction errors: {stats.MeshExtractionErrors}");
            sb.AppendLine($"- Mesh point tests: {stats.MeshPointTests}");
            sb.AppendLine($"- Bounds point tests: {stats.BoundsPointTests}");
            sb.AppendLine($"- Mesh fallback point tests: {stats.MeshFallbackPointTests}");
            sb.AppendLine();

            if (settings.ZoneContainmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate)
            {
                AppendMeshAccurateTierSummary(sb, settings);
            }

            if (settings.ZoneContainmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate)
            {
                var fallbackZones = dataset?.Zones?
                    .Where(z => !string.IsNullOrWhiteSpace(z.MeshFallbackReason))
                    .ToList() ?? new List<ZoneGeometry>();

                sb.AppendLine("### Mesh Fallback Reasons (mesh unusable)");
                sb.AppendLine($"- Fallback zones: {fallbackZones.Count}");
                if (fallbackZones.Count == 0)
                {
                    sb.AppendLine("- <none>");
                    sb.AppendLine();
                }
                else
                {
                    foreach (var group in fallbackZones.GroupBy(z => z.MeshFallbackReason))
                    {
                        sb.AppendLine($"- {group.Key}: {group.Count()}");
                    }

                    sb.AppendLine();
                    sb.AppendLine("#### Fallback Details (first 50)");
                    foreach (var zone in fallbackZones.Take(50))
                    {
                        var name = string.IsNullOrWhiteSpace(zone.DisplayName) ? zone.ZoneId : zone.DisplayName;
                        var detail = string.IsNullOrWhiteSpace(zone.MeshFallbackDetail) ? "<none>" : zone.MeshFallbackDetail;
                        sb.AppendLine($"- {name} ({zone.ZoneId}): {zone.MeshFallbackReason} - {detail}");
                    }
                    sb.AppendLine();
                }
            }

            if (settings.ZoneContainmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate)
            {
                sb.AppendLine("### Mesh Closure Warnings (mesh used)");
                sb.AppendLine($"- Non-watertight zones: {meshClosureWarnings.Count}");
                if (meshClosureWarnings.Count == 0)
                {
                    sb.AppendLine("- <none>");
                    sb.AppendLine();
                }
                else
                {
                    sb.AppendLine("#### Non-watertight Details (first 50)");
                    foreach (var zone in meshClosureWarnings.Take(50))
                    {
                        var name = string.IsNullOrWhiteSpace(zone.DisplayName) ? zone.ZoneId : zone.DisplayName;
                        sb.AppendLine($"- {name} ({zone.ZoneId}): BoundaryEdges={zone.MeshBoundaryEdgeCount}, NonManifoldEdges={zone.MeshNonManifoldEdgeCount}");
                    }
                    sb.AppendLine();
                }
            }

            AppendGpuDiagnostics(sb, stats);
            AppendGpuZoneDiagnostics(sb, stats);
            AppendSlowZoneDiagnostics(sb, stats);

            sb.AppendLine("## Step 4 - Outputs (Mappings)");
            var mappings = request?.Mappings ?? new List<SpaceMapperMappingDefinition>();
            for (int i = 0; i < mappings.Count; i++)
            {
                var mapping = mappings[i];
                sb.AppendLine($"- {i + 1}. {mapping.Name}: {mapping.ZoneCategory}.{mapping.ZonePropertyName} -> {mapping.TargetCategory}.{mapping.TargetPropertyName}, WriteMode={mapping.WriteMode}, AppendSeparator=\"{mapping.AppendSeparator}\", PartialFlag=\"{mapping.PartialFlagValue}\", MultiZone={mapping.MultiZoneCombineMode}");
            }
            if (mappings.Count == 0)
            {
                sb.AppendLine("- <none>");
            }
            sb.AppendLine();

            sb.AppendLine("## Step 5 - Results");
            sb.AppendLine($"- Status: {result?.Message ?? "<no message>"}");
            sb.AppendLine($"- Zones processed: {stats.ZonesProcessed}");
            sb.AppendLine($"- Targets processed: {stats.TargetsProcessed}");
            sb.AppendLine($"- Intersections: {result?.Intersections?.Count ?? 0}");
            sb.AppendLine($"- Contained tagged: {stats.ContainedTagged}");
            sb.AppendLine($"- Partial tagged: {stats.PartialTagged}");
            sb.AppendLine($"- Multi-zone tagged: {stats.MultiZoneTagged}");
            sb.AppendLine($"- Skipped: {stats.Skipped}");
            sb.AppendLine($"- Skipped unchanged: {stats.SkippedUnchanged}");
            sb.AppendLine($"- Writes performed: {stats.WritesPerformed}");
            sb.AppendLine($"- Targets written: {stats.WritebackTargetsWritten}");
            sb.AppendLine($"- Categories written: {stats.WritebackCategoriesWritten}");
            sb.AppendLine($"- Properties written: {stats.WritebackPropertiesWritten}");
            sb.AppendLine($"- Writeback strategy used: {stats.WritebackStrategy}");
            sb.AppendLine($"- Mode used: {stats.ModeUsed}");
            sb.AppendLine($"- Preset used: {stats.PresetUsed}");
            sb.AppendLine($"- Traversal used: {stats.TraversalUsed}");
            sb.AppendLine($"- Candidate pairs: {stats.CandidatePairs}");
            sb.AppendLine($"- Avg candidates/zone: {stats.AvgCandidatesPerZone:0.##}");
            sb.AppendLine($"- Max candidates/zone: {stats.MaxCandidatesPerZone}");
            sb.AppendLine($"- Avg candidates/target: {stats.AvgCandidatesPerTarget:0.##}");
            sb.AppendLine($"- Max candidates/target: {stats.MaxCandidatesPerTarget}");
            sb.AppendLine($"- Used preflight index: {stats.UsedPreflightIndex}");
            sb.AppendLine($"- Timings (ms): resolve {stats.ResolveTime.TotalMilliseconds:0}, build {stats.BuildGeometryTime.TotalMilliseconds:0}, index {stats.BuildIndexTime.TotalMilliseconds:0}, candidates {stats.CandidateQueryTime.TotalMilliseconds:0}, narrow {stats.NarrowPhaseTime.TotalMilliseconds:0}, write {stats.WriteBackTime.TotalMilliseconds:0}, total {stats.Elapsed.TotalMilliseconds:0}");
            sb.AppendLine();

            if (stats.ZoneSummaries != null && stats.ZoneSummaries.Count > 0)
            {
                sb.AppendLine("### Zone Summaries");
                foreach (var zone in stats.ZoneSummaries.OrderByDescending(z => z.ContainedCount + z.PartialCount).Take(50))
                {
                    sb.AppendLine($"- {zone.ZoneName} ({zone.ZoneId}): contained={zone.ContainedCount}, partial={zone.PartialCount}");
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private static string BuildLiveReport(
            SpaceMapperLiveReportCache cache,
            SpaceMapperRequest request,
            SpaceMapperResolvedData resolved,
            SpaceMapperComputeDataset dataset,
            SpaceMapperRunResult result,
            SpaceMapperRunProgressState progress)
        {
            var stats = result?.Stats ?? new SpaceMapperRunStats();
            var settings = request?.ProcessingSettings ?? new SpaceMapperProcessingSettings();
            var nowLocal = DateTime.Now;
            var nowUtc = DateTime.UtcNow;

            using var proc = Process.GetCurrentProcess();
            var sb = new StringBuilder();

            sb.AppendLine("# Space Mapper Live Report");
            sb.AppendLine("Auto-updated snapshot (this file is overwritten during the run).");
            sb.AppendLine($"Generated (UTC): {nowUtc:u}");
            sb.AppendLine($"Generated (Local): {nowLocal:yyyy-MM-dd HH:mm:ss}");
            if (cache != null)
            {
                sb.AppendLine($"Run started (UTC): {cache.RunStartUtc:u}");
            }
            sb.AppendLine();

            sb.AppendLine("## Environment");
            sb.AppendLine($"- Model: {cache?.ModelTitle ?? "<unknown>"}");
            sb.AppendLine($"- Model path: {cache?.ModelPath ?? "<unsaved>"}");
            sb.AppendLine($"- Navisworks API: {typeof(Application).Assembly.GetName().Version}");
            sb.AppendLine($"- Plugin version: {typeof(SpaceMapperService).Assembly.GetName().Version}");
            sb.AppendLine($"- OS: {Environment.OSVersion}");
            sb.AppendLine($"- Machine: {Environment.MachineName}");
            sb.AppendLine($"- CPU cores: {Environment.ProcessorCount}");
            sb.AppendLine($"- Working set (MB): {proc.WorkingSet64 / (1024.0 * 1024.0):0.0}");
            sb.AppendLine($"- Managed heap (MB): {GC.GetTotalMemory(false) / (1024.0 * 1024.0):0.0}");
            sb.AppendLine($"- Thread count: {proc.Threads.Count}");
            try
            {
                var startUtc = proc.StartTime.ToUniversalTime();
                sb.AppendLine($"- Process ID: {proc.Id}");
                sb.AppendLine($"- Process start (UTC): {startUtc:u}");
                sb.AppendLine($"- Process uptime (min): {(DateTime.UtcNow - startUtc).TotalMinutes:0.0}");
            }
            catch
            {
                sb.AppendLine("- Process info: <unavailable>");
            }
            sb.AppendLine($"- Is 64-bit process: {Environment.Is64BitProcess}");
            sb.AppendLine($"- CLR version: {Environment.Version}");
            sb.AppendLine($"- Base directory: {AppDomain.CurrentDomain.BaseDirectory}");
            sb.AppendLine();

            sb.AppendLine("## Graphics Environment");
            AppendD3D11AdapterInfo(sb, stats);
            if (!string.IsNullOrWhiteSpace(cache?.GraphicsSnapshot))
            {
                sb.AppendLine(cache.GraphicsSnapshot.TrimEnd());
                sb.AppendLine();
            }

            AppendProgressSnapshot(sb, progress);

            sb.AppendLine("## Step 1 - Setup");
            sb.AppendLine($"- Template: {request?.TemplateName ?? "<none>"}");
            sb.AppendLine($"- Scraper profile: {request?.ScraperProfileName ?? "<none>"}");
            sb.AppendLine($"- Scope: {request?.Scope}");
            sb.AppendLine($"- Zone source: {request?.ZoneSource}");
            sb.AppendLine($"- Zone set/search: {request?.ZoneSetName ?? "<none>"}");
            sb.AppendLine($"- Target rule count: {request?.TargetRules?.Count ?? 0}");
            sb.AppendLine($"- Mapping count: {request?.Mappings?.Count ?? 0}");
            sb.AppendLine();

            sb.AppendLine("## Step 2 - Zones & Targets");
            sb.AppendLine($"- Resolved zones (raw): {resolved?.ZoneModels?.Count ?? 0}");
            sb.AppendLine($"- Zones with bbox: {dataset?.Zones?.Count ?? 0}");
            sb.AppendLine($"- Resolved targets (raw): {resolved?.TargetModels?.Count ?? 0}");
            sb.AppendLine($"- Targets with bbox: {dataset?.TargetsForEngine?.Count ?? 0}");
            sb.AppendLine($"- Exclude zones from targets: {settings.ExcludeZonesFromTargets}");
            sb.AppendLine();

            sb.AppendLine("## Step 3 - Processing");
            sb.AppendLine($"- Processing mode: {settings.ProcessingMode}");
            sb.AppendLine($"- GPU ray count: {Math.Max(settings.GpuRayCount, 1)}");
            sb.AppendLine($"- Preset: {settings.PerformancePreset}");
            sb.AppendLine($"- Fast traversal: {settings.FastTraversalMode}");
            sb.AppendLine($"- Zone containment engine: {settings.ZoneContainmentEngine}");
            sb.AppendLine($"- Zone bounds: {settings.ZoneBoundsMode} (k-DOP: {settings.ZoneKDopVariant})");
            sb.AppendLine($"- Target bounds: {settings.TargetBoundsMode} (k-DOP: {settings.TargetKDopVariant}, midpoint: {settings.TargetMidpointMode})");
            sb.AppendLine($"- Use origin point only: {settings.UseOriginPointOnly}");
            sb.AppendLine($"- Resolution strategy: {settings.ZoneResolutionStrategy}");
            sb.AppendLine($"- Enable multiple zones: {settings.EnableMultipleZones}");
            sb.AppendLine($"- Treat partial as contained: {settings.TreatPartialAsContained}");
            sb.AppendLine($"- Tag partial separately: {settings.TagPartialSeparately}");
            sb.AppendLine($"- Write zone behavior property: {settings.WriteZoneBehaviorProperty}");
            sb.AppendLine($"- Write containment % property: {settings.WriteZoneContainmentPercentProperty}");
            sb.AppendLine($"- Containment calculation mode: {settings.ContainmentCalculationMode}");
            sb.AppendLine($"- Zone behavior category: {settings.ZoneBehaviorCategory}");
            sb.AppendLine($"- Zone behavior property: {settings.ZoneBehaviorPropertyName}");
            sb.AppendLine($"- Zone behavior contained value: {settings.ZoneBehaviorContainedValue}");
            sb.AppendLine($"- Zone behavior partial value: {settings.ZoneBehaviorPartialValue}");
            sb.AppendLine($"- Offset 3D: {settings.Offset3D}");
            sb.AppendLine($"- Offset sides: {settings.OffsetSides}");
            sb.AppendLine($"- Offset top: {settings.OffsetTop}");
            sb.AppendLine($"- Offset bottom: {settings.OffsetBottom}");
            sb.AppendLine($"- Units: {settings.Units}");
            sb.AppendLine($"- Offset mode: {settings.OffsetMode}");
            sb.AppendLine($"- Index granularity: {settings.IndexGranularity}");
            sb.AppendLine($"- Max threads: {settings.MaxThreads}");
            sb.AppendLine($"- Batch size: {settings.BatchSize}");
            sb.AppendLine($"- Writeback strategy: {settings.WritebackStrategy}");
            sb.AppendLine($"- Show internal properties during writeback: {settings.ShowInternalPropertiesDuringWriteback}");
            sb.AppendLine($"- Skip unchanged writeback: {settings.SkipUnchangedWriteback}");
            sb.AppendLine($"- Pack writeback outputs: {settings.PackWritebackProperties}");
            sb.AppendLine($"- Close dock panes during run: {settings.CloseDockPanesDuringRun}");
            sb.AppendLine();

            sb.AppendLine("## Live Results");
            sb.AppendLine($"- Status: {result?.Message ?? "<running>"}");
            if (progress != null)
            {
                sb.AppendLine($"- Zones processed (progress): {progress.ZonesProcessed}/{progress.ZonesTotal}");
                sb.AppendLine($"- Targets processed (progress): {progress.TargetsProcessed}/{progress.TargetsTotal}");
                sb.AppendLine($"- Writeback targets (progress): {progress.WriteTargetsProcessed}/{progress.WriteTargetsTotal}");
                sb.AppendLine($"- Candidate pairs (progress): {progress.CandidatePairs:N0}");
                sb.AppendLine($"- Elapsed (progress): {progress.Elapsed:hh\\:mm\\:ss}");
            }
            sb.AppendLine($"- Intersections: {result?.Intersections?.Count ?? 0}");
            sb.AppendLine($"- Contained tagged: {stats.ContainedTagged}");
            sb.AppendLine($"- Partial tagged: {stats.PartialTagged}");
            sb.AppendLine($"- Multi-zone tagged: {stats.MultiZoneTagged}");
            sb.AppendLine($"- Writes performed: {stats.WritesPerformed}");
            sb.AppendLine($"- Skipped unchanged: {stats.SkippedUnchanged}");
            sb.AppendLine($"- Mode used: {stats.ModeUsed}");
            sb.AppendLine($"- Preset used: {stats.PresetUsed}");
            sb.AppendLine($"- Traversal used: {stats.TraversalUsed}");
            sb.AppendLine($"- Timings (ms): resolve {stats.ResolveTime.TotalMilliseconds:0}, build {stats.BuildGeometryTime.TotalMilliseconds:0}, index {stats.BuildIndexTime.TotalMilliseconds:0}, candidates {stats.CandidateQueryTime.TotalMilliseconds:0}, narrow {stats.NarrowPhaseTime.TotalMilliseconds:0}, write {stats.WriteBackTime.TotalMilliseconds:0}, total {stats.Elapsed.TotalMilliseconds:0}");
            sb.AppendLine();

            AppendGpuDiagnostics(sb, stats);
            AppendSlowZoneDiagnostics(sb, stats);

            return sb.ToString();
        }

        private static void AppendProgressSnapshot(StringBuilder sb, SpaceMapperRunProgressState progress)
        {
            if (progress == null)
            {
                return;
            }

            sb.AppendLine("## Run Progress Snapshot");
            sb.AppendLine($"- Stage: {progress.Stage} ({progress.StageText})");
            sb.AppendLine($"- Detail: {(string.IsNullOrWhiteSpace(progress.DetailText) ? "<none>" : progress.DetailText)}");
            sb.AppendLine($"- Elapsed: {progress.Elapsed:hh\\:mm\\:ss}");
            sb.AppendLine($"- Zones: {progress.ZonesProcessed}/{progress.ZonesTotal}");
            sb.AppendLine($"- Targets: {progress.TargetsProcessed}/{progress.TargetsTotal}");
            sb.AppendLine($"- Writeback: {progress.WriteTargetsProcessed}/{progress.WriteTargetsTotal}");
            sb.AppendLine($"- Candidate pairs (progress): {progress.CandidatePairs:N0}");
            sb.AppendLine($"- Stage started (UTC): {progress.StageStartUtc:u}");
            sb.AppendLine($"- Last progress (UTC): {progress.LastProgressUtc:u}");
            sb.AppendLine($"- Last zone update (UTC): {progress.LastZoneProgressUtc:u}");
            sb.AppendLine($"- Last target update (UTC): {progress.LastTargetProgressUtc:u}");
            sb.AppendLine($"- Last writeback update (UTC): {progress.LastWriteProgressUtc:u}");
            if (progress.IsCancelled)
            {
                sb.AppendLine("- Status: Cancelled");
            }
            if (progress.IsFailed)
            {
                sb.AppendLine($"- Status: Failed ({progress.ErrorText})");
            }
            sb.AppendLine();

            var timeline = progress.GetTimelineSnapshot();
            if (timeline != null && timeline.Count > 0)
            {
                sb.AppendLine("### Stage Timeline");
                foreach (var entry in timeline)
                {
                    var endUtc = entry.EndUtc?.ToString("u") ?? "<running>";
                    var detail = string.IsNullOrWhiteSpace(entry.DetailText) ? "<none>" : entry.DetailText;
                    sb.AppendLine($"- {entry.Stage} ({entry.StageText})");
                    sb.AppendLine($"  - Start (UTC): {entry.StartUtc:u}");
                    sb.AppendLine($"  - End (UTC): {endUtc}");
                    sb.AppendLine($"  - Detail: {detail}");
                }
                sb.AppendLine();
            }
        }

        private static void AppendSlowZoneDiagnostics(StringBuilder sb, SpaceMapperRunStats stats)
        {
            sb.AppendLine("### Slow Zone Diagnostics");
            var threshold = stats?.SlowZoneThresholdSeconds ?? 0;
            sb.AppendLine($"- Threshold (sec): {(threshold > 0 ? threshold.ToString("0.##", CultureInfo.InvariantCulture) : "<unknown>")}");

            var slowZones = stats?.SlowZones ?? new List<SpaceMapperSlowZoneInfo>();
            if (slowZones.Count == 0)
            {
                sb.AppendLine("- <none>");
                sb.AppendLine();
                return;
            }

            var max = slowZones.Count > 50 ? 50 : slowZones.Count;
            foreach (var zone in slowZones.OrderByDescending(z => z.Elapsed).Take(max))
            {
                var name = string.IsNullOrWhiteSpace(zone.ZoneName) ? zone.ZoneId : zone.ZoneName;
                sb.AppendLine($"- {name} ({zone.ZoneId}) index {zone.ZoneIndex + 1}: {zone.Elapsed.TotalSeconds:0.##}s, candidates={zone.CandidateCount}");
            }
            sb.AppendLine();
        }

        private static void AppendGpuDiagnostics(StringBuilder sb, SpaceMapperRunStats stats)
        {
            if (stats == null)
            {
                return;
            }

            sb.AppendLine("### GPU Diagnostics");
            sb.AppendLine($"- Backend: {stats.GpuBackend ?? "<none>"}");
            if (!string.IsNullOrWhiteSpace(stats.GpuInitFailureReason))
            {
                sb.AppendLine($"- Init failure: {stats.GpuInitFailureReason}");
            }
            sb.AppendLine($"- Zones processed on GPU: {stats.GpuZonesProcessed}");
            sb.AppendLine($"- Points tested: {stats.GpuPointsTested}");
            sb.AppendLine($"- Triangles tested: {stats.GpuTrianglesTested}");
            if (stats.GpuZonesProcessed > 0)
            {
                var avg = stats.GpuTrianglesTested / (double)stats.GpuZonesProcessed;
                sb.AppendLine($"- Avg triangles/zone: {avg:0.##}");
            }
            if (stats.GpuPointThreshold > 0)
            {
                sb.AppendLine($"- GPU point threshold: {stats.GpuPointThreshold}");
            }
            if (stats.GpuSamplePointsPerTarget > 0)
            {
                sb.AppendLine($"- Sample points/target: {stats.GpuSamplePointsPerTarget}");
            }
            sb.AppendLine($"- Zones eligible for GPU: {stats.GpuZonesEligible}");
            sb.AppendLine($"- Zones skipped (no mesh): {stats.GpuZonesSkippedNoMesh}");
            sb.AppendLine($"- Zones skipped (missing triangles): {stats.GpuZonesSkippedMissingTriangles}");
            sb.AppendLine($"- Zones skipped (open mesh): {stats.GpuZonesSkippedOpenMesh}");
            sb.AppendLine($"- Zones skipped (below point threshold): {stats.GpuZonesSkippedLowPoints}");
            sb.AppendLine($"- Uncertain points (CPU fallback): {stats.GpuUncertainPoints}");
            if (stats.GpuOpenMeshBoundaryEdgeLimit > 0 || stats.GpuOpenMeshNonManifoldEdgeLimit > 0)
            {
                sb.AppendLine($"- Open mesh tolerance: boundary<= {stats.GpuOpenMeshBoundaryEdgeLimit}, nonmanifold<= {stats.GpuOpenMeshNonManifoldEdgeLimit}, outsideSamples<= {stats.GpuOpenMeshOutsideTolerance}");
                sb.AppendLine($"- Open mesh zones eligible: {stats.GpuOpenMeshZonesEligible}");
                sb.AppendLine($"- Open mesh zones processed on GPU: {stats.GpuOpenMeshZonesProcessed}");
                if (stats.GpuOpenMeshNudge > 0)
                {
                    sb.AppendLine($"- Open mesh point nudge: {stats.GpuOpenMeshNudge:0.######}");
                }
            }
            if (stats.GpuMaxTrianglesPerZone > 0)
            {
                sb.AppendLine($"- Max triangles/zone: {stats.GpuMaxTrianglesPerZone}");
            }
            if (stats.GpuMaxPointsPerZone > 0)
            {
                sb.AppendLine($"- Max points/zone: {stats.GpuMaxPointsPerZone}");
            }
            if (stats.GpuBatchDispatchCount > 0)
            {
                sb.AppendLine($"- Batch dispatches: {stats.GpuBatchDispatchCount}");
                sb.AppendLine($"- Max zones/dispatch: {stats.GpuBatchMaxZones}");
                sb.AppendLine($"- Max points/dispatch: {stats.GpuBatchMaxPoints}");
                sb.AppendLine($"- Max triangles/dispatch: {stats.GpuBatchMaxTriangles}");
                sb.AppendLine($"- Avg zones/dispatch: {stats.GpuBatchAvgZonesPerDispatch:0.##}");
            }
            sb.AppendLine($"- GPU dispatch time (ms): {stats.GpuDispatchTime.TotalMilliseconds:0}");
            sb.AppendLine($"- GPU readback time (ms): {stats.GpuReadbackTime.TotalMilliseconds:0}");
            sb.AppendLine();
        }

        private static void AppendGpuZoneDiagnostics(StringBuilder sb, SpaceMapperRunStats stats)
        {
            if (stats?.GpuZoneDiagnostics == null || stats.GpuZoneDiagnostics.Count == 0)
            {
                return;
            }

            sb.AppendLine("#### GPU Zone Eligibility");
            sb.AppendLine("| Zone | Candidates | EstPoints | Triangles | Work | PointThreshold | WorkThreshold | Eligible | UseGpu | Packed | OpenMesh | SkipReason |");
            sb.AppendLine("| --- | ---: | ---: | ---: | ---: | ---: | ---: | --- | --- | --- | --- | --- |");
            foreach (var zone in stats.GpuZoneDiagnostics)
            {
                var name = string.IsNullOrWhiteSpace(zone.ZoneName) ? zone.ZoneId : zone.ZoneName;
                var openMesh = zone.IsOpenMesh
                    ? (zone.AllowOpenMeshGpu ? "open-ok" : "open-skip")
                    : "closed";
                var skip = string.IsNullOrWhiteSpace(zone.SkipReason) ? "-" : zone.SkipReason;
                sb.AppendLine($"| {EscapeMarkdownCell(name)} | {zone.CandidateCount} | {zone.EstimatedPoints} | {zone.TriangleCount} | {zone.WorkEstimate} | {zone.PointThreshold} | {zone.WorkThreshold} | {(zone.EligibleForGpu ? "yes" : "no")} | {(zone.UsedGpu ? "yes" : "no")} | {(zone.PackedThresholds ? "yes" : "no")} | {openMesh} | {EscapeMarkdownCell(skip)} |");
            }
            sb.AppendLine();
        }

        private static void AppendGraphicsEnvironment(StringBuilder sb, SpaceMapperRunStats stats)
        {
            sb.AppendLine("## Graphics Environment");
            AppendD3D11AdapterInfo(sb, stats);
            AppendVideoControllerInfo(sb);
            AppendGraphicsDriverRegistry(sb);
            sb.AppendLine();
        }

        private static string BuildGraphicsSnapshot()
        {
            var sb = new StringBuilder();
            AppendVideoControllerInfo(sb);
            AppendGraphicsDriverRegistry(sb);
            return sb.ToString().TrimEnd();
        }

        private static void AppendD3D11AdapterInfo(StringBuilder sb, SpaceMapperRunStats stats)
        {
            sb.AppendLine("### D3D11 Adapter (GPU compute)");
            if (stats == null)
            {
                sb.AppendLine("- <unavailable>");
                sb.AppendLine();
                return;
            }

            var hasAdapter = !string.IsNullOrWhiteSpace(stats.GpuAdapterName)
                || stats.GpuVendorId.HasValue
                || stats.GpuDeviceId.HasValue;
            if (!hasAdapter)
            {
                sb.AppendLine("- <not initialized>");
                sb.AppendLine();
                return;
            }

            sb.AppendLine($"- Name: {stats.GpuAdapterName ?? "<unknown>"}");
            if (stats.GpuVendorId.HasValue || stats.GpuDeviceId.HasValue || stats.GpuSubSysId.HasValue || stats.GpuRevision.HasValue)
            {
                sb.AppendLine($"- IDs: vendor={FormatHex(stats.GpuVendorId)}, device={FormatHex(stats.GpuDeviceId)}, subsys={FormatHex(stats.GpuSubSysId)}, rev={FormatHex(stats.GpuRevision)}");
            }
            if (!string.IsNullOrWhiteSpace(stats.GpuAdapterLuid))
            {
                sb.AppendLine($"- LUID: {stats.GpuAdapterLuid}");
            }
            if (!string.IsNullOrWhiteSpace(stats.GpuFeatureLevel))
            {
                sb.AppendLine($"- Feature level: {stats.GpuFeatureLevel}");
            }
            if (stats.GpuDedicatedVideoMemory.HasValue)
            {
                sb.AppendLine($"- Dedicated video memory: {FormatBytes(stats.GpuDedicatedVideoMemory.Value)}");
            }
            if (stats.GpuDedicatedSystemMemory.HasValue)
            {
                sb.AppendLine($"- Dedicated system memory: {FormatBytes(stats.GpuDedicatedSystemMemory.Value)}");
            }
            if (stats.GpuSharedSystemMemory.HasValue)
            {
                sb.AppendLine($"- Shared system memory: {FormatBytes(stats.GpuSharedSystemMemory.Value)}");
            }
            sb.AppendLine();
        }

        private static void AppendVideoControllerInfo(StringBuilder sb)
        {
            sb.AppendLine("### Video Controllers (WMI)");
            try
            {
                using var searcher = new ManagementObjectSearcher(
                    "SELECT Name, DriverVersion, DriverDate, AdapterRAM, PNPDeviceID, VideoProcessor, Status FROM Win32_VideoController");
                using var results = searcher.Get();
                var any = false;
                foreach (ManagementObject item in results)
                {
                    any = true;
                    var name = item["Name"]?.ToString() ?? "<unknown>";
                    sb.AppendLine($"- {name}");
                    sb.AppendLine($"  - Driver version: {item["DriverVersion"] ?? "<unknown>"}");
                    sb.AppendLine($"  - Driver date: {FormatWmiDate(item["DriverDate"])}");
                    sb.AppendLine($"  - Adapter RAM: {FormatBytes(AsInt64(item["AdapterRAM"]))}");
                    sb.AppendLine($"  - PNP device ID: {item["PNPDeviceID"] ?? "<unknown>"}");
                    sb.AppendLine($"  - Video processor: {item["VideoProcessor"] ?? "<unknown>"}");
                    sb.AppendLine($"  - Status: {item["Status"] ?? "<unknown>"}");
                }
                if (!any)
                {
                    sb.AppendLine("- <none>");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"- <error reading WMI: {ex.Message}>");
            }
            sb.AppendLine();
        }

        private static void AppendGraphicsDriverRegistry(StringBuilder sb)
        {
            sb.AppendLine("### Graphics Drivers Registry");
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\GraphicsDrivers");
                if (key == null)
                {
                    sb.AppendLine("- <not available>");
                    sb.AppendLine();
                    return;
                }

                sb.AppendLine($"- HwSchMode: {FormatHwSchMode(key.GetValue("HwSchMode"))}");
                sb.AppendLine($"- TdrDelay: {FormatRegistryValue(key.GetValue("TdrDelay"))}");
                sb.AppendLine($"- TdrDdiDelay: {FormatRegistryValue(key.GetValue("TdrDdiDelay"))}");
                sb.AppendLine($"- TdrLevel: {FormatTdrLevel(key.GetValue("TdrLevel"))}");
                sb.AppendLine($"- TdrLimitTime: {FormatRegistryValue(key.GetValue("TdrLimitTime"))}");
                sb.AppendLine($"- TdrLimitCount: {FormatRegistryValue(key.GetValue("TdrLimitCount"))}");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"- <error reading registry: {ex.Message}>");
            }
            sb.AppendLine();
        }

        private static void AppendTargetTransformSection(
            StringBuilder sb,
            SpaceMapperComputeDataset dataset,
            SpaceMapperProcessingSettings settings,
            ISet<string> intersectionTargets)
        {
            var targets = dataset?.TargetsForEngine ?? new List<TargetGeometry>();
            var summary = new TargetTransformSummary();
            sb.AppendLine("### Target Transforms & Midpoints");
            sb.AppendLine($"- Targets listed: {targets.Count}");

            if (targets.Count == 0)
            {
                sb.AppendLine("- <none>");
                sb.AppendLine();
                return;
            }

            var includeInternal = settings?.ShowInternalPropertiesDuringWriteback == true;
            var dumpedPropertyKeys = false;
            var limit = targets.Count > 200 ? 200 : targets.Count;
            for (int i = 0; i < limit; i++)
            {
                var target = targets[i];
                if (target == null)
                {
                    continue;
                }

                var name = string.IsNullOrWhiteSpace(target.DisplayName) ? target.ItemKey : target.DisplayName;
                sb.AppendLine($"- {name} ({target.ItemKey})");

                var hasPoint = AppendBoundingBoxProperties(sb, target.BoundingBox, settings, out var pointUsed);
                summary.TargetsTotal++;
                if (hasPoint)
                {
                    summary.TargetsWithPoint++;
                }

                var fragmentStats = AppendFragmentTransformProperties(sb, target.ModelItem, hasPoint, pointUsed, ref dumpedPropertyKeys);
                UpdateTransformSummary(summary, fragmentStats, hasPoint);

                var hasIntersection = intersectionTargets != null
                    && !string.IsNullOrWhiteSpace(target.ItemKey)
                    && intersectionTargets.Contains(target.ItemKey);
                if (hasIntersection)
                {
                    summary.TargetsWithIntersections++;
                    if (hasPoint && fragmentStats?.MinDistanceToPoint.HasValue == true)
                    {
                        if (fragmentStats.MinDistanceToPoint.Value <= PointMatchTolerance)
                        {
                            summary.TargetsPointMatchesWithIntersections++;
                        }
                        else
                        {
                            summary.TargetsPointDiffersWithIntersections++;
                        }
                    }
                    else
                    {
                        summary.TargetsPointUnknownWithIntersections++;
                    }
                }
                if (includeInternal)
                {
                    AppendInternalTransformProperties(sb, target.ModelItem);
                }
            }

            if (targets.Count > limit)
            {
                sb.AppendLine($"- ... truncated {targets.Count - limit} targets");
            }

            AppendTransformSummary(sb, summary);
            sb.AppendLine();
        }

        private static void AppendMeshAccurateTierSummary(StringBuilder sb, SpaceMapperProcessingSettings settings)
        {
            if (sb == null || settings == null)
            {
                return;
            }

            var effectiveTargetMode = GetEffectiveMeshAccurateTargetMode(settings);
            var isTierA = effectiveTargetMode == SpaceMapperTargetBoundsMode.Midpoint;
            var tierLabel = isTierA
                ? "Tier A (Midpoint: single point)"
                : $"Tier B ({effectiveTargetMode}: multi-sample)";

            sb.AppendLine("### Mesh Accurate Tier");
            sb.AppendLine($"- Tier: {tierLabel}");
            if (effectiveTargetMode != settings.TargetBoundsMode)
            {
                sb.AppendLine($"- Effective target bounds: {effectiveTargetMode} (UseOriginPointOnly)");
            }
            sb.AppendLine($"- Samples per target: {GetMeshAccurateSampleCount(effectiveTargetMode)}");
            sb.AppendLine();
        }

        private static int GetMeshAccurateSampleCount(SpaceMapperTargetBoundsMode targetBoundsMode)
        {
            return targetBoundsMode switch
            {
                SpaceMapperTargetBoundsMode.Aabb => 8,
                SpaceMapperTargetBoundsMode.Obb => 9,
                SpaceMapperTargetBoundsMode.KDop => 15,
                SpaceMapperTargetBoundsMode.Hull => 27,
                _ => 1
            };
        }

        private static SpaceMapperTargetBoundsMode GetEffectiveMeshAccurateTargetMode(SpaceMapperProcessingSettings settings)
        {
            if (settings == null)
            {
                return SpaceMapperTargetBoundsMode.Aabb;
            }

            if (settings.TargetBoundsMode == SpaceMapperTargetBoundsMode.Aabb && settings.UseOriginPointOnly)
            {
                return SpaceMapperTargetBoundsMode.Midpoint;
            }

            return settings.TargetBoundsMode;
        }

        private static bool AppendBoundingBoxProperties(
            StringBuilder sb,
            BoundingBox3D bbox,
            SpaceMapperProcessingSettings settings,
            out Point3D pointUsed)
        {
            pointUsed = new Point3D();
            if (bbox == null)
            {
                sb.AppendLine("  - Bounds: <bbox missing>");
                return false;
            }

            var min = bbox.Min;
            var max = bbox.Max;
            var center = new Point3D(
                (min.X + max.X) * 0.5,
                (min.Y + max.Y) * 0.5,
                (min.Z + max.Z) * 0.5);
            var bottomCenter = new Point3D(
                (min.X + max.X) * 0.5,
                (min.Y + max.Y) * 0.5,
                min.Z);

            sb.AppendLine($"  - BBox min: {FormatPoint(min)}");
            sb.AppendLine($"  - BBox max: {FormatPoint(max)}");
            sb.AppendLine($"  - BBox center: {FormatPoint(center)}");
            sb.AppendLine($"  - BBox bottom center: {FormatPoint(bottomCenter)}");

            if (settings != null)
            {
                var useBottom = settings.TargetMidpointMode == SpaceMapperMidpointMode.BoundingBoxBottomCenter;
                var usePoint = settings.TargetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint
                    || settings.UseOriginPointOnly;
                if (usePoint)
                {
                    pointUsed = useBottom ? bottomCenter : center;
                    sb.AppendLine($"  - Point used ({settings.TargetMidpointMode}): {FormatPoint(pointUsed)}");
                    return true;
                }
            }

            return false;
        }

        private static FragmentTransformStats AppendFragmentTransformProperties(
            StringBuilder sb,
            ModelItem item,
            bool hasPointUsed,
            Point3D pointUsed,
            ref bool dumpedPropertyKeys)
        {
            var stats = new FragmentTransformStats();
            var transforms = TryGetFragmentTransforms(item, maxFragments: 5, out var error);
            if (transforms == null)
            {
                sb.AppendLine("  - Fragments: <unavailable>");
                if (!string.IsNullOrWhiteSpace(error))
                {
                    sb.AppendLine($"  - Fragment extraction error: {error}");
                }

                if (!dumpedPropertyKeys)
                {
                    AppendPropertyKeyDump(sb, item, 40);
                    dumpedPropertyKeys = true;
                }

                return stats;
            }

            stats.HasTransforms = true;
            stats.FragmentCount = transforms.FragmentCount;
            stats.DistinctTranslationCount = transforms.DistinctTranslationCount;

            sb.AppendLine($"  - Fragments: {transforms.FragmentCount}");
            if (transforms.FragmentCount == 0)
            {
                if (!string.IsNullOrWhiteSpace(error))
                {
                    sb.AppendLine($"  - Fragment extraction error: {error}");
                }

                if (!dumpedPropertyKeys)
                {
                    AppendPropertyKeyDump(sb, item, 40);
                    dumpedPropertyKeys = true;
                }

                return stats;
            }

            if (transforms.Matrices.Count == 0)
            {
                sb.AppendLine("  - Fragment matrices: <missing>");
            }
            else
            {
                var translationSamples = 0;
                double? minDistance = null;

                for (var i = 0; i < transforms.Matrices.Count; i++)
                {
                    if (i < transforms.TranslationSamples.Count && transforms.TranslationSamples[i] != null)
                    {
                        var info = transforms.TranslationSamples[i];
                        translationSamples++;
                        sb.AppendLine($"  - Fragment[{i}] Translation ({info.Source}): {FormatPoint(info.Translation)}");
                        if (hasPointUsed)
                        {
                            AppendPointDelta(sb, pointUsed, info.Translation, "(point - fragment)");
                        }
                        if (info.HasAlternate)
                        {
                            sb.AppendLine($"  - Fragment[{i}] Translation (alt): {FormatPoint(info.Alternate)}");
                            if (hasPointUsed)
                            {
                                AppendPointDelta(sb, pointUsed, info.Alternate, "(point - fragment alt)");
                            }
                        }
                    }
                    sb.AppendLine($"  - Fragment[{i}] Matrix: {FormatMatrix(transforms.Matrices[i])}");

                    if (hasPointUsed && i < transforms.TranslationSamples.Count && transforms.TranslationSamples[i] != null)
                    {
                        var info = transforms.TranslationSamples[i];
                        var dist = Distance(pointUsed, info.Translation);
                        minDistance = minDistance.HasValue ? Math.Min(minDistance.Value, dist) : dist;
                    }
                }

                stats.TranslationSampleCount = translationSamples;
                if (hasPointUsed)
                {
                    stats.MinDistanceToPoint = minDistance;
                }
            }

            if (transforms.FragmentCount > 1)
            {
                if (transforms.TranslationCount == 0)
                {
                    sb.AppendLine("  - Distinct fragment translations: <missing>");
                }
                else if (transforms.DistinctTranslationCount <= 1)
                {
                    sb.AppendLine("  - Distinct fragment translations: 1 (all same)");
                }
                else
                {
                    sb.AppendLine($"  - Distinct fragment translations: {transforms.DistinctTranslationCount}");
                }
            }

            if (transforms.Matrices.Count == 0)
            {
                stats.TranslationSampleCount = 0;
            }

            return stats;
        }

        private static void AppendInternalTransformProperties(StringBuilder sb, ModelItem item)
        {
            var scaleX = ReadPropertyValue(item, "Transform", "Scale.X");
            var scaleY = ReadPropertyValue(item, "Transform", "Scale.Y");
            var scaleZ = ReadPropertyValue(item, "Transform", "Scale.Z");
            var rotationAngle = ReadPropertyValue(item, "Transform", "Rotation Angle");
            var rotationAxisX = ReadPropertyValue(item, "Transform", "Rotation Axis.X");
            var rotationAxisY = ReadPropertyValue(item, "Transform", "Rotation Axis.Y");
            var rotationAxisZ = ReadPropertyValue(item, "Transform", "Rotation Axis.Z");
            var translationX = ReadPropertyValue(item, "Transform", "Translation.X");
            var translationY = ReadPropertyValue(item, "Transform", "Translation.Y");
            var translationZ = ReadPropertyValue(item, "Transform", "Translation.Z");
            var reverses = ReadPropertyValue(item, "Transform", "Reverses");

            var hasAny = !string.IsNullOrWhiteSpace(scaleX)
                || !string.IsNullOrWhiteSpace(scaleY)
                || !string.IsNullOrWhiteSpace(scaleZ)
                || !string.IsNullOrWhiteSpace(rotationAngle)
                || !string.IsNullOrWhiteSpace(rotationAxisX)
                || !string.IsNullOrWhiteSpace(rotationAxisY)
                || !string.IsNullOrWhiteSpace(rotationAxisZ)
                || !string.IsNullOrWhiteSpace(translationX)
                || !string.IsNullOrWhiteSpace(translationY)
                || !string.IsNullOrWhiteSpace(translationZ)
                || !string.IsNullOrWhiteSpace(reverses);

            if (!hasAny)
            {
                sb.AppendLine("  - Internal properties (best effort): <missing>");
                return;
            }

            sb.AppendLine(
                $"  - Internal properties (best effort): Scale=({FormatValue(scaleX)},{FormatValue(scaleY)},{FormatValue(scaleZ)}), " +
                $"RotationAngle={FormatValue(rotationAngle)}, " +
                $"RotationAxis=({FormatValue(rotationAxisX)},{FormatValue(rotationAxisY)},{FormatValue(rotationAxisZ)}), " +
                $"Translation=({FormatValue(translationX)},{FormatValue(translationY)},{FormatValue(translationZ)}), " +
                $"Reverses={FormatValue(reverses)}");
        }

        private sealed class FragmentTransformReport
        {
            public int FragmentCount { get; set; }
            public int TranslationCount { get; set; }
            public int DistinctTranslationCount { get; set; }
            public List<double[]> Matrices { get; } = new();
            public List<FragmentTranslationInfo> TranslationSamples { get; } = new();
        }

        private sealed class FragmentTranslationInfo
        {
            public Point3D Translation { get; set; }
            public string Source { get; set; }
            public bool HasAlternate { get; set; }
            public Point3D Alternate { get; set; }
        }

        private sealed class FragmentTransformStats
        {
            public bool HasTransforms { get; set; }
            public int FragmentCount { get; set; }
            public int TranslationSampleCount { get; set; }
            public int DistinctTranslationCount { get; set; }
            public double? MinDistanceToPoint { get; set; }
        }

        private sealed class TargetTransformSummary
        {
            public int TargetsTotal { get; set; }
            public int TargetsWithPoint { get; set; }
            public int TargetsWithFragments { get; set; }
            public int TargetsWithTranslationSamples { get; set; }
            public int TargetsMissingFragments { get; set; }
            public int TargetsPointMatchesFragment { get; set; }
            public int TargetsPointDiffersFragment { get; set; }
            public int TargetsPointNoTranslation { get; set; }
            public int TargetsMultipleDistinctTranslations { get; set; }
            public int TargetsWithIntersections { get; set; }
            public int TargetsPointMatchesWithIntersections { get; set; }
            public int TargetsPointDiffersWithIntersections { get; set; }
            public int TargetsPointUnknownWithIntersections { get; set; }
            public double MinDistanceSum { get; set; }
            public double MinDistanceMax { get; set; }
            public int MinDistanceCount { get; set; }
        }

        private static void UpdateTransformSummary(
            TargetTransformSummary summary,
            FragmentTransformStats stats,
            bool hasPointUsed)
        {
            if (summary == null)
            {
                return;
            }

            if (stats == null || !stats.HasTransforms)
            {
                summary.TargetsMissingFragments++;
                if (hasPointUsed)
                {
                    summary.TargetsPointNoTranslation++;
                }
                return;
            }

            if (stats.FragmentCount > 0)
            {
                summary.TargetsWithFragments++;
            }
            else
            {
                summary.TargetsMissingFragments++;
            }

            if (stats.TranslationSampleCount > 0)
            {
                summary.TargetsWithTranslationSamples++;
            }
            else if (hasPointUsed)
            {
                summary.TargetsPointNoTranslation++;
            }

            if (stats.DistinctTranslationCount > 1)
            {
                summary.TargetsMultipleDistinctTranslations++;
            }

            if (hasPointUsed && stats.MinDistanceToPoint.HasValue)
            {
                summary.MinDistanceCount++;
                summary.MinDistanceSum += stats.MinDistanceToPoint.Value;
                summary.MinDistanceMax = Math.Max(summary.MinDistanceMax, stats.MinDistanceToPoint.Value);
                if (stats.MinDistanceToPoint.Value <= PointMatchTolerance)
                {
                    summary.TargetsPointMatchesFragment++;
                }
                else
                {
                    summary.TargetsPointDiffersFragment++;
                }
            }
        }

        private static void AppendTransformSummary(StringBuilder sb, TargetTransformSummary summary)
        {
            if (summary == null || summary.TargetsTotal == 0)
            {
                return;
            }

            sb.AppendLine("#### Transform Comparison Summary");
            sb.AppendLine($"- Targets processed: {summary.TargetsTotal}");
            sb.AppendLine($"- Targets with point used: {summary.TargetsWithPoint}");
            sb.AppendLine($"- Targets with fragment translations (sampled): {summary.TargetsWithTranslationSamples}");
            sb.AppendLine($"- Point matches fragment (<= {FormatScalar(PointMatchTolerance)}): {summary.TargetsPointMatchesFragment}");
            sb.AppendLine($"- Point differs from fragment: {summary.TargetsPointDiffersFragment}");
            if (summary.TargetsPointNoTranslation > 0)
            {
                sb.AppendLine($"- Point used but no fragment translation: {summary.TargetsPointNoTranslation}");
            }
            if (summary.TargetsMultipleDistinctTranslations > 0)
            {
                sb.AppendLine($"- Targets with multiple distinct fragment translations: {summary.TargetsMultipleDistinctTranslations}");
            }
            if (summary.MinDistanceCount > 0)
            {
                var avg = summary.MinDistanceSum / summary.MinDistanceCount;
                sb.AppendLine($"- Min distance avg: {FormatScalar(avg)}, max: {FormatScalar(summary.MinDistanceMax)}");
            }
            if (summary.TargetsWithIntersections > 0)
            {
                sb.AppendLine($"- Targets with intersections: {summary.TargetsWithIntersections}");
                sb.AppendLine($"- Intersections with point match: {summary.TargetsPointMatchesWithIntersections}");
                sb.AppendLine($"- Intersections with point differ: {summary.TargetsPointDiffersWithIntersections}");
                if (summary.TargetsPointUnknownWithIntersections > 0)
                {
                    sb.AppendLine($"- Intersections with unknown point delta: {summary.TargetsPointUnknownWithIntersections}");
                }
            }
        }

        private static FragmentTransformReport TryGetFragmentTransforms(
            ModelItem item,
            int maxFragments,
            out string error)
        {
            error = null;
            if (item == null)
            {
                return null;
            }

            var dispatcher = System.Windows.Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
            {
                FragmentTransformReport report = null;
                string localError = null;
                dispatcher.Invoke(new Action(() =>
                {
                    report = TryGetFragmentTransformsOnUiThread(item, maxFragments, out localError);
                }));
                error = localError;
                return report;
            }

            return TryGetFragmentTransformsOnUiThread(item, maxFragments, out error);
        }

        private static FragmentTransformReport TryGetFragmentTransformsOnUiThread(
            ModelItem item,
            int maxFragments,
            out string error)
        {
            error = null;
            try
            {
                var selection = ComBridge.ToInwOpSelection(new ModelItemCollection { item });
                var paths = selection?.Paths();
                if (paths == null)
                {
                    return null;
                }

                var report = new FragmentTransformReport();
                var distinctTranslations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (ComApi.InwOaPath3 path in paths)
                {
                    if (path == null)
                    {
                        continue;
                    }

                    var fragments = path.Fragments();
                    if (fragments == null)
                    {
                        continue;
                    }

                    foreach (ComApi.InwOaFragment3 fragment in fragments)
                    {
                        if (fragment == null)
                        {
                            continue;
                        }

                        report.FragmentCount++;

                        double[] matrix = null;
                        try
                        {
                            matrix = ExtractMatrix(fragment.GetLocalToWorldMatrix()?.Matrix);
                        }
                        catch
                        {
                            matrix = null;
                        }

                        if (matrix == null || matrix.Length < 12)
                        {
                            continue;
                        }

                        var hasTranslation = TryGetTranslation(matrix, out var translation, out var source, out var alternate, out var hasAlternate);
                        if (hasTranslation)
                        {
                            report.TranslationCount++;
                            distinctTranslations.Add(FormatPoint(translation));
                        }

                        if (report.Matrices.Count < maxFragments)
                        {
                            report.Matrices.Add(matrix);
                            if (hasTranslation)
                            {
                                report.TranslationSamples.Add(new FragmentTranslationInfo
                                {
                                    Translation = translation,
                                    Source = source,
                                    HasAlternate = hasAlternate,
                                    Alternate = alternate
                                });
                            }
                            else
                            {
                                report.TranslationSamples.Add(null);
                            }
                        }
                    }
                }

                report.DistinctTranslationCount = distinctTranslations.Count;
                return report;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return null;
            }
        }

        private static double[] ExtractMatrix(object matrixObj)
        {
            if (matrixObj is Array arr && arr.Length >= 12)
            {
                var start = arr.GetLowerBound(0);
                var result = new double[arr.Length];
                for (int i = 0; i < arr.Length; i++)
                {
                    result[i] = ToDouble(arr.GetValue(start + i));
                }
                return result;
            }

            return null;
        }

        private static bool TryGetTranslation(
            double[] matrix,
            out Point3D translation,
            out string source,
            out Point3D alternate,
            out bool hasAlternate)
        {
            translation = new Point3D();
            source = "matrix";
            alternate = new Point3D();
            hasAlternate = false;
            if (matrix == null || matrix.Length < 12)
            {
                return false;
            }

            if (matrix.Length >= 16)
            {
                var row = new Point3D(matrix[12], matrix[13], matrix[14]);
                var col = new Point3D(matrix[3], matrix[7], matrix[11]);
                var rowZero = IsNearZero(row);
                var colZero = IsNearZero(col);

                if (!rowZero && colZero)
                {
                    translation = row;
                    source = "row4";
                    return true;
                }

                if (!colZero && rowZero)
                {
                    translation = col;
                    source = "col4";
                    return true;
                }

                if (!rowZero && !colZero)
                {
                    translation = row;
                    source = "row4";
                    if (!PointsClose(row, col))
                    {
                        alternate = col;
                        hasAlternate = true;
                    }
                    return true;
                }

                translation = row;
                source = "row4";
                return true;
            }

            translation = new Point3D(matrix[3], matrix[7], matrix[11]);
            source = "col4";
            return true;
        }

        private static bool IsNearZero(Point3D point)
        {
            return Math.Abs(point.X) < 1e-6
                && Math.Abs(point.Y) < 1e-6
                && Math.Abs(point.Z) < 1e-6;
        }

        private static bool PointsClose(Point3D left, Point3D right)
        {
            return Math.Abs(left.X - right.X) < 1e-6
                && Math.Abs(left.Y - right.Y) < 1e-6
                && Math.Abs(left.Z - right.Z) < 1e-6;
        }

        private static void AppendPointDelta(StringBuilder sb, Point3D pointUsed, Point3D fragmentTranslation, string label)
        {
            var dx = pointUsed.X - fragmentTranslation.X;
            var dy = pointUsed.Y - fragmentTranslation.Y;
            var dz = pointUsed.Z - fragmentTranslation.Z;
            var dist = Math.Sqrt(dx * dx + dy * dy + dz * dz);
            var delta = new Point3D(dx, dy, dz);
            sb.AppendLine(
                $"  - Delta to point used {label}: {FormatPoint(delta)}, dist={FormatScalar(dist)}");
        }

        private static string FormatMatrix(double[] matrix)
        {
            if (matrix == null || matrix.Length < 12)
            {
                return "<missing>";
            }

            var rows = matrix.Length >= 16 ? 4 : 3;
            var sb = new StringBuilder();
            sb.Append("[");
            for (int row = 0; row < rows; row++)
            {
                if (row > 0)
                {
                    sb.Append(" ");
                }
                sb.Append("[");
                var rowStart = row * 4;
                for (int col = 0; col < 4; col++)
                {
                    if (col > 0)
                    {
                        sb.Append(", ");
                    }
                    var value = rowStart + col < matrix.Length ? matrix[rowStart + col] : 0.0;
                    sb.AppendFormat(CultureInfo.InvariantCulture, "{0:0.###}", value);
                }
                sb.Append("]");
                if (row < rows - 1)
                {
                    sb.Append(",");
                }
            }
            sb.Append("]");
            return sb.ToString();
        }

        private static void AppendPropertyKeyDump(StringBuilder sb, ModelItem item, int limit)
        {
            if (item == null)
            {
                return;
            }

            sb.AppendLine("  - Internal properties (keys only, first target):");
            var count = 0;
            try
            {
                foreach (var category in item.PropertyCategories)
                {
                    if (category == null)
                    {
                        continue;
                    }

                    var catName = category.DisplayName ?? category.Name ?? "<unnamed>";
                    foreach (var prop in category.Properties)
                    {
                        if (prop == null)
                        {
                            continue;
                        }

                        var propName = prop.DisplayName ?? prop.Name ?? "<unnamed>";
                        sb.AppendLine($"    - {catName} | {propName}");
                        count++;
                        if (count >= limit)
                        {
                            sb.AppendLine("    - ... truncated");
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"    - <error reading property keys: {ex.Message}>");
            }
        }

        private static double ToDouble(object value)
        {
            return value == null ? 0.0 : Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        private static string BuildReportPath(Document doc, string templateName)
        {
            var baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "Autodesk",
                "Navisworks Manage 2025",
                "Plugins",
                "MicroEng.Navisworks",
                "Reports",
                "SpaceMapper");

            var modelName = GetDocumentTitle(doc);
            var template = string.IsNullOrWhiteSpace(templateName) ? null : templateName.Trim();
            var safeModel = SanitizeFileName(modelName);
            var safeTemplate = SanitizeFileName(template);
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            var suffixParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(safeModel)) suffixParts.Add(safeModel);
            if (!string.IsNullOrWhiteSpace(safeTemplate)) suffixParts.Add(safeTemplate);
            var suffix = suffixParts.Count == 0 ? string.Empty : "_" + string.Join("_", suffixParts);

            var fileName = $"SpaceMapper_Run_{stamp}{suffix}.md";
            return Path.Combine(baseDir, fileName);
        }

        private static string BuildLiveReportPath(Document doc, string templateName)
        {
            var baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "Autodesk",
                "Navisworks Manage 2025",
                "Plugins",
                "MicroEng.Navisworks",
                "Reports",
                "SpaceMapper");

            var modelName = GetDocumentTitle(doc);
            var template = string.IsNullOrWhiteSpace(templateName) ? null : templateName.Trim();
            var safeModel = SanitizeFileName(modelName);
            var safeTemplate = SanitizeFileName(template);
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            var suffixParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(safeModel)) suffixParts.Add(safeModel);
            if (!string.IsNullOrWhiteSpace(safeTemplate)) suffixParts.Add(safeTemplate);
            var suffix = suffixParts.Count == 0 ? string.Empty : "_" + string.Join("_", suffixParts);

            var fileName = $"SpaceMapper_Live_{stamp}{suffix}.md";
            return Path.Combine(baseDir, fileName);
        }

        private static string GetDocumentTitle(Document doc)
        {
            if (doc == null)
            {
                return "Untitled";
            }

            var title = doc.Title;
            if (!string.IsNullOrWhiteSpace(title))
            {
                return title.Trim();
            }

            if (!string.IsNullOrWhiteSpace(doc.FileName))
            {
                return Path.GetFileNameWithoutExtension(doc.FileName);
            }

            return "Untitled";
        }

        private static string SanitizeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var invalid = Path.GetInvalidFileNameChars();
            var cleaned = new string(value.Where(ch => !invalid.Contains(ch)).ToArray());
            return cleaned.Length > 60 ? cleaned.Substring(0, 60) : cleaned;
        }

        private static string FormatRange(int? min, int? max)
        {
            if (min == null && max == null)
            {
                return "<any>";
            }

            if (min != null && max != null)
            {
                return $"{min}-{max}";
            }

            if (min != null)
            {
                return $">= {min}";
            }

            return $"<= {max}";
        }

        private static string ReadPropertyValue(ModelItem item, string categoryName, string propertyName)
        {
            if (item == null || string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(propertyName))
            {
                return string.Empty;
            }

            return ReadPropertyValue(item, categoryName, propertyName, true);
        }

        private static string ReadPropertyValue(
            ModelItem item,
            string categoryName,
            string propertyName,
            bool allowCategoryFallback)
        {
            if (item == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return string.Empty;
            }

            var requireCategory = !string.IsNullOrWhiteSpace(categoryName);
            var targetCategoryKey = NormalizeKey(categoryName);
            var targetPropertyKey = NormalizeKey(propertyName);

            try
            {
                foreach (var cat in item.PropertyCategories)
                {
                    var catKey = NormalizeKey(cat.DisplayName ?? cat.Name);
                    if (requireCategory && !CategoryMatches(catKey, targetCategoryKey))
                    {
                        continue;
                    }

                    foreach (var prop in cat.Properties)
                    {
                        var displayKey = NormalizeKey(prop.DisplayName);
                        var nameKey = NormalizeKey(prop.Name);
                        if (PropertyMatches(targetPropertyKey, displayKey)
                            || PropertyMatches(targetPropertyKey, nameKey))
                        {
                            return prop.Value?.ToDisplayString() ?? prop.Value?.ToString() ?? string.Empty;
                        }
                    }
                }
            }
            catch
            {
                // ignore read failures
            }

            if (requireCategory && allowCategoryFallback)
            {
                return ReadPropertyValue(item, null, propertyName, false);
            }

            return string.Empty;
        }

        private static bool PropertyMatches(string targetKey, string candidateKey)
        {
            if (string.IsNullOrWhiteSpace(targetKey) || string.IsNullOrWhiteSpace(candidateKey))
            {
                return false;
            }

            if (string.Equals(candidateKey, targetKey, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (candidateKey.IndexOf(targetKey, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (TransformPropertyAliases.TryGetValue(targetKey, out var aliases))
            {
                foreach (var alias in aliases)
                {
                    if (candidateKey.IndexOf(alias, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool CategoryMatches(string categoryKey, string targetKey)
        {
            if (string.IsNullOrWhiteSpace(targetKey))
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(categoryKey))
            {
                return false;
            }

            if (string.Equals(categoryKey, targetKey, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return categoryKey.IndexOf(targetKey, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string NormalizeKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = NormalizeDisplayName(value);
            var builder = new StringBuilder(normalized.Length);
            foreach (var ch in normalized)
            {
                if (char.IsLetterOrDigit(ch))
                {
                    builder.Append(char.ToLowerInvariant(ch));
                }
            }

            return builder.ToString();
        }

        private static string NormalizeDisplayName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var trimmed = value.Trim();
            var suffixIndex = trimmed.LastIndexOf(" (", StringComparison.Ordinal);
            if (suffixIndex > 0 && trimmed.EndsWith(")", StringComparison.Ordinal))
            {
                return trimmed.Substring(0, suffixIndex);
            }

            return trimmed;
        }

        private static string FormatValue(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "<missing>" : value;
        }

        private static string FormatPoint(Point3D point)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "({0:0.###}, {1:0.###}, {2:0.###})",
                point.X,
                point.Y,
                point.Z);
        }

        private static string FormatScalar(double value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static string EscapeMarkdownCell(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "-";
            }

            return value.Replace("|", "\\|");
        }

        private static string FormatHex(int? value)
        {
            return value.HasValue ? $"0x{value.Value:x}" : "<unknown>";
        }

        private static string FormatBytes(long bytes)
        {
            if (bytes < 0)
            {
                return "<unknown>";
            }

            const double scale = 1024.0;
            if (bytes >= scale * scale * scale)
            {
                return $"{bytes / (scale * scale * scale):0.##} GB";
            }
            if (bytes >= scale * scale)
            {
                return $"{bytes / (scale * scale):0.##} MB";
            }
            if (bytes >= scale)
            {
                return $"{bytes / scale:0.##} KB";
            }
            return $"{bytes} B";
        }

        private static long AsInt64(object value)
        {
            if (value == null)
            {
                return -1;
            }

            try
            {
                return Convert.ToInt64(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return -1;
            }
        }

        private static string FormatWmiDate(object value)
        {
            if (value == null)
            {
                return "<unknown>";
            }

            var raw = value.ToString();
            if (string.IsNullOrWhiteSpace(raw))
            {
                return "<unknown>";
            }

            try
            {
                var dt = ManagementDateTimeConverter.ToDateTime(raw);
                return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }
            catch
            {
                return raw;
            }
        }

        private static string FormatRegistryValue(object value)
        {
            if (value == null)
            {
                return "<not set>";
            }

            try
            {
                return Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                return value.ToString();
            }
        }

        private static string FormatHwSchMode(object value)
        {
            if (value == null)
            {
                return "<not set>";
            }

            int mode;
            try
            {
                mode = Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return value.ToString();
            }

            var label = mode switch
            {
                0 => "Default",
                1 => "Disabled",
                2 => "Enabled",
                _ => "Unknown"
            };
            return $"{mode} ({label})";
        }

        private static string FormatTdrLevel(object value)
        {
            if (value == null)
            {
                return "<not set>";
            }

            int level;
            try
            {
                level = Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return value.ToString();
            }

            var label = level switch
            {
                0 => "Off",
                1 => "Bugcheck",
                2 => "RecoverVGA",
                3 => "Recover",
                _ => "Unknown"
            };
            return $"{level} ({label})";
        }

        private static double Distance(Point3D left, Point3D right)
        {
            var dx = left.X - right.X;
            var dy = left.Y - right.Y;
            var dz = left.Z - right.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        private static ISet<string> BuildIntersectionTargetSet(IEnumerable<ZoneTargetIntersection> intersections)
        {
            if (intersections == null)
            {
                return null;
            }

            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var inter in intersections)
            {
                if (inter == null || string.IsNullOrWhiteSpace(inter.TargetItemKey))
                {
                    continue;
                }
                set.Add(inter.TargetItemKey);
            }

            return set;
        }

        private static Dictionary<SpaceMapperTargetRule, int> BuildRuleCounts(
            SpaceMapperResolvedData resolved,
            IReadOnlyList<SpaceMapperTargetRule> rules)
        {
            var counts = new Dictionary<SpaceMapperTargetRule, int>(new ReferenceEqualityComparer<SpaceMapperTargetRule>());
            if (rules != null)
            {
                foreach (var rule in rules)
                {
                    if (rule == null)
                    {
                        continue;
                    }
                    counts[rule] = 0;
                }
            }

            if (resolved?.TargetsByRule == null)
            {
                return counts;
            }

            foreach (var entry in resolved.TargetsByRule.Values)
            {
                if (entry == null) continue;
                foreach (var rule in entry)
                {
                    if (rule == null) continue;
                    if (!counts.ContainsKey(rule))
                    {
                        counts[rule] = 0;
                    }
                    counts[rule]++;
                }
            }

            return counts;
        }

        private sealed class ReferenceEqualityComparer<T> : IEqualityComparer<T> where T : class
        {
            public bool Equals(T x, T y) => ReferenceEquals(x, y);
            public int GetHashCode(T obj) => obj == null ? 0 : RuntimeHelpers.GetHashCode(obj);
        }
    }
}
