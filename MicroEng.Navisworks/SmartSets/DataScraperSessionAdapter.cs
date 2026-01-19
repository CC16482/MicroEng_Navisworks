using System;
using System.Collections.Generic;
using System.Linq;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.SmartSets
{
    internal sealed class DataScraperSessionAdapter : IDataScrapeSessionView
    {
        private readonly ScrapeSession _session;

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
                foreach (var p in _session.Properties ?? Enumerable.Empty<ScrapedProperty>())
                {
                    var samples = p.SampleValues?.Select(s => s ?? "").ToList() ?? new List<string>();
                    yield return new ScrapedPropertyDescriptor(
                        p.Category ?? "",
                        p.Name ?? "",
                        p.DataType ?? "",
                        p.ItemCount,
                        p.DistinctValueCount,
                        samples);
                }
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
