using System;
using System.Collections.Generic;
using System.Linq;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.SmartSets
{
    internal sealed class DataScraperSessionAdapter : IDataScrapeSessionView
    {
        private readonly ScrapeSession _session;
        private IReadOnlyList<ScrapedPropertyDescriptor> _cachedProperties;

        public DataScraperSessionAdapter(ScrapeSession session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
        }

        public Guid SessionId => _session.Id;

        public DateTime Timestamp => _session.Timestamp;

        public string ProfileName => _session.ProfileName ?? "";

        public int ItemsScanned => _session.ItemsScanned;

        public IEnumerable<ScrapedPropertyDescriptor> Properties
        {
            get
            {
                if (_cachedProperties != null)
                {
                    return _cachedProperties;
                }

                var list = new List<ScrapedPropertyDescriptor>();
                foreach (var p in _session.Properties ?? Enumerable.Empty<ScrapedProperty>())
                {
                    var samples = p.SampleValues?.Count > 0
                        ? p.SampleValues.Select(s => s ?? string.Empty).ToList()
                        : new List<string>();
                    list.Add(new ScrapedPropertyDescriptor(
                        p.Category ?? string.Empty,
                        p.Name ?? string.Empty,
                        p.DataType ?? string.Empty,
                        p.ItemCount,
                        p.DistinctValueCount,
                        samples));
                }

                _cachedProperties = list;
                return _cachedProperties;
            }
        }

        public IEnumerable<ScrapedRawEntryView> RawEntries
        {
            get
            {
                foreach (var r in _session.RawEntries ?? Enumerable.Empty<RawEntry>())
                {
                    yield return new ScrapedRawEntryView
                    {
                        ItemPath = r.ItemPath ?? "",
                        Category = r.Category ?? "",
                        Property = r.Name ?? "",
                        Value = r.Value ?? ""
                    };
                }
            }
        }
    }
}
