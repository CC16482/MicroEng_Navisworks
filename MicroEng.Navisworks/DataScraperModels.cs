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
