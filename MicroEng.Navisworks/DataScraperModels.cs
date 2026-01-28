using System;
using System.Collections.Generic;
using System.Linq;

namespace MicroEng.Navisworks
{
    internal class ScrapedProperty
    {
        public string Category { get; set; }
        public string Name { get; set; }
        public string DataType { get; set; }
        public int ItemCount { get; set; }
        public int DistinctValueCount { get; set; }
        public List<string> SampleValues { get; set; } = new();
    }

    internal class ScrapeSession
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public string ProfileName { get; set; }
        public string ScopeType { get; set; }
        public string ScopeDescription { get; set; }
        public int ItemsScanned { get; set; }
        public string DocumentFile { get; set; }
        public string DocumentFileKey { get; set; }

        public bool JsonlExportEnabled { get; set; }
        public string JsonlExportPath { get; set; }
        public string JsonlExportMode { get; set; }
        public bool JsonlExportGzip { get; set; }
        public string JsonlPrimaryKeyMode { get; set; }
        public long JsonlLinesWritten { get; set; }
        public bool JsonlStreamedDuringScrape { get; set; }

        public bool RawEntriesTruncated { get; set; }
        public List<ScrapedProperty> Properties { get; set; } = new();
        public List<RawEntry> RawEntries { get; set; } = new();
    }

    internal class RawEntry
    {
        public string Profile { get; set; }
        public string Scope { get; set; }
        public string ItemKey { get; set; }
        public string ItemPath { get; set; }
        public string Category { get; set; }
        public string Name { get; set; }
        public string DataType { get; set; }
        public string Value { get; set; }
    }

    internal enum DataScraperJsonlExportMode
    {
        RawRows = 0,
        ItemDocuments = 1
    }

    internal enum DataScraperExportSource
    {
        PreviewRaw = 0,
        FullRaw = 1,
        PropertySummaries = 2
    }

    internal enum DataScraperPrimaryKeyMode
    {
        InstanceGuid = 0,
        BestExternalId = 1,
        SourceFilePlusBestExternalId = 2,
        ItemPath = 3
    }

    internal enum DataScraperPathKeyMode
    {
        FileNameOnly = 0,
        FullPath = 1
    }

    internal sealed class DataScraperJsonlExportSettings
    {
        public bool Enabled { get; set; }
        public bool StreamDuringScrape { get; set; } = true;

        public DataScraperJsonlExportMode Mode { get; set; } = DataScraperJsonlExportMode.RawRows;
        public DataScraperExportSource ExportSource { get; set; } = DataScraperExportSource.PreviewRaw;
        public bool ExportRawRows { get; set; } = true;
        public bool ExportItemDocuments { get; set; }
        public bool Gzip { get; set; }
        public string OutputPath { get; set; } = "";

        public DataScraperPrimaryKeyMode PrimaryKeyMode { get; set; } = DataScraperPrimaryKeyMode.SourceFilePlusBestExternalId;
        public DataScraperPathKeyMode SourceFileKeyMode { get; set; } = DataScraperPathKeyMode.FileNameOnly;

        public bool IncludeDocumentFile { get; set; } = true;
        public bool IncludeSourceFile { get; set; } = true;

        public bool IncludeNavisworksInstanceGuid { get; set; } = true;
        public bool IncludeRevitIds { get; set; } = true;
        public bool IncludeIfcIds { get; set; } = true;
        public bool IncludeDwgIds { get; set; } = true;

        public bool NormalizeIds { get; set; } = true;

        public bool KeepRawEntriesInMemory { get; set; } = true;
        public int PreviewRawRowLimit { get; set; } = 50000;
    }

    internal static class DataScraperCache
    {
        public static ScrapeSession LastSession { get; set; }
        public static List<ScrapeSession> AllSessions { get; } = new();
        public static event Action<ScrapeSession> SessionAdded;

        public static IEnumerable<string> GetAllPropertyNames()
        {
            return LastSession?.Properties.Select(p => p.Name).Distinct(StringComparer.OrdinalIgnoreCase) ??
                   Enumerable.Empty<string>();
        }

        public static IEnumerable<(string Category, string Name)> GetCategoryAndNames()
        {
            return LastSession?.Properties.Select(p => (p.Category, p.Name)) ??
                   Enumerable.Empty<(string, string)>();
        }

        public static void AddSession(ScrapeSession session)
        {
            LastSession = session;
            AllSessions.Add(session);
            SessionAdded?.Invoke(session);
        }
    }
}
