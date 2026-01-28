using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using Autodesk.Navisworks.Api;
using NavisApp = Autodesk.Navisworks.Api.Application;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace MicroEng.Navisworks
{
    internal enum ScrapeScopeType
    {
        SingleItem,
        CurrentSelection,
        SelectionSet,
        SearchSet,
        EntireModel
    }

    internal class DataScraperService
    {
        public ScrapeSession Scrape(string profileName, ScrapeScopeType scopeType, string scopeDescription, IEnumerable<ModelItem> items)
        {
            return Scrape(profileName, scopeType, scopeDescription, items, export: null);
        }

        public ScrapeSession Scrape(
            string profileName,
            ScrapeScopeType scopeType,
            string scopeDescription,
            IEnumerable<ModelItem> items,
            DataScraperJsonlExportSettings export)
        {
            var utcNow = DateTime.UtcNow;
            var session = new ScrapeSession
            {
                ProfileName = profileName,
                ScopeType = scopeType.ToString(),
                ScopeDescription = scopeDescription,
                Timestamp = DateTime.Now
            };

            var doc = NavisApp.ActiveDocument;
            var documentFile = doc?.FileName ?? string.Empty;
            session.DocumentFile = documentFile;

            var exportEnabled = export != null && export.Enabled;
            var exportRawRows = exportEnabled && export.ExportRawRows;
            var exportItemDocs = exportEnabled && export.ExportItemDocuments;
            if (!exportRawRows && !exportItemDocs)
            {
                exportEnabled = exportEnabled && export.ExportSource == DataScraperExportSource.PropertySummaries;
            }
            var exportStreaming = exportEnabled && export.StreamDuringScrape;
            var keepAllRaw = export?.KeepRawEntriesInMemory ?? true;
            var previewLimit = keepAllRaw ? int.MaxValue : (export?.PreviewRawRowLimit ?? 0);
            if (!keepAllRaw && previewLimit < 0)
            {
                previewLimit = 0;
            }

            session.RawEntries = new List<RawEntry>();
            session.RawEntriesTruncated = !keepAllRaw && previewLimit == 0;

            session.DocumentFileKey = BuildPathKey(documentFile, export?.SourceFileKeyMode ?? DataScraperPathKeyMode.FileNameOnly);

            string rawExportPath = null;
            string itemExportPath = null;
            string summaryExportPath = null;

            if (exportEnabled)
            {
                session.JsonlExportEnabled = true;
                session.JsonlExportMode = exportRawRows && exportItemDocs
                    ? "RawRows+ItemDocuments"
                    : exportItemDocs
                        ? DataScraperJsonlExportMode.ItemDocuments.ToString()
                        : DataScraperJsonlExportMode.RawRows.ToString();
                if (export.ExportSource == DataScraperExportSource.PropertySummaries)
                {
                    session.JsonlExportMode = "PropertySummaries";
                }
                session.JsonlExportGzip = export.Gzip;
                session.JsonlPrimaryKeyMode = export.PrimaryKeyMode.ToString();
                session.JsonlStreamedDuringScrape = export.StreamDuringScrape;

                rawExportPath = exportRawRows
                    ? BuildExportPath(export.OutputPath, profileName, utcNow, export.Gzip, exportItemDocs ? "_raw" : string.Empty)
                    : null;
                itemExportPath = exportItemDocs
                    ? BuildExportPath(export.OutputPath, profileName, utcNow, export.Gzip, exportRawRows ? "_items" : string.Empty)
                    : null;
                summaryExportPath = export.ExportSource == DataScraperExportSource.PropertySummaries
                    ? BuildExportPath(export.OutputPath, profileName, utcNow, export.Gzip, "_summary")
                    : null;

                if (exportRawRows && exportItemDocs)
                {
                    session.JsonlExportPath = $"{rawExportPath} | {itemExportPath}";
                }
                else
                {
                    session.JsonlExportPath = rawExportPath ?? itemExportPath;
                    if (summaryExportPath != null)
                    {
                        session.JsonlExportPath = summaryExportPath;
                    }
                }
            }

            JsonlStreamWriter rawWriter = null;
            JsonlStreamWriter itemWriter = null;
            if (exportStreaming)
            {
                if (exportRawRows && !string.IsNullOrWhiteSpace(rawExportPath))
                {
                    rawWriter = new JsonlStreamWriter(rawExportPath, export.Gzip);
                }
                if (exportItemDocs && !string.IsNullOrWhiteSpace(itemExportPath))
                {
                    itemWriter = new JsonlStreamWriter(itemExportPath, export.Gzip);
                }
            }

            List<JsonlRawRowRecord> pendingRawRows = null;
            List<JsonlItemDocRecord> pendingItemDocs = null;
            if (exportEnabled && !exportStreaming)
            {
                if (exportRawRows)
                {
                    pendingRawRows = new List<JsonlRawRowRecord>();
                }
                if (exportItemDocs)
                {
                    pendingItemDocs = new List<JsonlItemDocRecord>();
                }
            }

            var propertyMap = new Dictionary<string, PropertyAccumulator>(StringComparer.OrdinalIgnoreCase);
            var sourceFileHit = 0;
            var revitUniqueIdHit = 0;
            var ifcGlobalIdHit = 0;
            var dwgHandleHit = 0;
            var rawEntriesAtLimit = previewLimit == 0;

            try
            {
                foreach (var item in items ?? Enumerable.Empty<ModelItem>())
                {
                    if (item == null)
                    {
                        continue;
                    }

                    session.ItemsScanned++;

                    var itemPath = ItemToPath(item);
                    var itemKey = item.InstanceGuid.ToString("D");

                    var probe = exportEnabled ? new IdentityProbe() : null;
                    var rawRowsForExport = exportRawRows ? new List<ItemRawRow>() : null;
                    var docProperties = exportItemDocs ? new List<JsonlItemDocProperty>() : null;
                    var seenPropsThisItem = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var category in item.PropertyCategories)
                    {
                        if (category == null)
                        {
                            continue;
                        }

                        var catName = category.DisplayName ?? category.Name ?? string.Empty;
                        var catInternal = category.Name ?? string.Empty;

                        foreach (var prop in category.Properties)
                        {
                            if (prop == null)
                            {
                                continue;
                            }

                            var name = prop.DisplayName ?? prop.Name ?? string.Empty;
                            var dtype = prop.Value?.DataType.ToString() ?? "Unknown";
                            var values = GetPropertyValueStrings(prop);
                            if (values.Count == 0)
                            {
                                continue;
                            }

                            values = values
                                .Where(v => !string.IsNullOrWhiteSpace(v))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList();
                            if (values.Count == 0)
                            {
                                continue;
                            }

                            probe?.Observe(catName, catInternal, name, prop.Name, values);

                            var key = $"{catName}\u001F{name}\u001F{dtype}";
                            var firstThisItem = seenPropsThisItem.Add(key);
                            TouchProperty(propertyMap, catName, name, dtype, firstThisItem, values);

                            if (!rawEntriesAtLimit)
                            {
                                foreach (var v in values)
                                {
                                    if (session.RawEntries.Count >= previewLimit)
                                    {
                                        session.RawEntriesTruncated = true;
                                        rawEntriesAtLimit = true;
                                        break;
                                    }

                                    session.RawEntries.Add(new RawEntry
                                    {
                                        Profile = profileName,
                                        Scope = scopeDescription,
                                        ItemKey = itemKey,
                                        ItemPath = itemPath,
                                        Category = catName,
                                        Name = name,
                                        DataType = dtype,
                                        Value = v
                                    });
                                }
                            }

                            if (rawRowsForExport != null)
                            {
                                foreach (var v in values)
                                {
                                    rawRowsForExport.Add(new ItemRawRow
                                    {
                                        Category = catName,
                                        Property = name,
                                        DataType = dtype,
                                        Value = v
                                    });
                                }
                            }

                            if (docProperties != null)
                            {
                                docProperties.Add(new JsonlItemDocProperty
                                {
                                    Category = catName,
                                    Property = name,
                                    DataType = dtype,
                                    Values = values
                                });
                            }
                        }
                    }

                    if (!exportEnabled)
                    {
                        continue;
                    }

                    var identity = BuildIdentity(item, itemPath, documentFile, probe, export);

                    if (!string.IsNullOrWhiteSpace(identity.SourceFile))
                    {
                        sourceFileHit++;
                    }
                    if (!string.IsNullOrWhiteSpace(identity.RevitUniqueId))
                    {
                        revitUniqueIdHit++;
                    }
                    if (!string.IsNullOrWhiteSpace(identity.IfcGlobalId))
                    {
                        ifcGlobalIdHit++;
                    }
                    if (!string.IsNullOrWhiteSpace(identity.DwgHandle))
                    {
                        dwgHandleHit++;
                    }

                    if (exportRawRows && rawRowsForExport != null)
                    {
                        foreach (var row in rawRowsForExport)
                        {
                            var parsed = ParseValue(row.Value);

                            var record = new JsonlRawRowRecord
                            {
                                SessionId = session.Id,
                                SessionTimestampUtc = utcNow,

                                ProfileName = profileName,
                                ScopeType = scopeType.ToString(),
                                ScopeDescription = scopeDescription,

                                DocumentFile = export.IncludeDocumentFile ? identity.DocumentFile : null,
                                DocumentFileKey = export.IncludeDocumentFile ? identity.DocumentFileKey : null,
                                SourceFile = export.IncludeSourceFile ? identity.SourceFile : null,
                                SourceFileKey = export.IncludeSourceFile ? identity.SourceFileKey : null,

                                PrimaryKey = identity.PrimaryKey,
                                PrimaryKeyType = identity.PrimaryKeyType,

                                NavisInstanceGuid = export.IncludeNavisworksInstanceGuid ? identity.NavisInstanceGuid : null,
                                RevitUniqueId = export.IncludeRevitIds ? identity.RevitUniqueId : null,
                                RevitElementId = export.IncludeRevitIds ? identity.RevitElementId : null,
                                IfcGlobalId = export.IncludeIfcIds ? identity.IfcGlobalId : null,
                                DwgHandle = export.IncludeDwgIds ? identity.DwgHandle : null,

                                Ids = identity.Ids,
                                ItemPath = itemPath,

                                Category = row.Category,
                                Property = row.Property,
                                DataType = row.DataType,
                                Value = row.Value,
                                ValueNorm = parsed.ValueNorm,
                                ValueNum = parsed.ValueNum,
                                ValueBool = parsed.ValueBool,
                                ValueDateUtc = parsed.ValueDateUtc
                            };

                            if (rawWriter != null)
                            {
                                rawWriter.Write(record);
                            }
                            else
                            {
                                pendingRawRows?.Add(record);
                            }
                        }
                    }
                    else if (exportItemDocs && docProperties != null)
                    {
                        var record = new JsonlItemDocRecord
                        {
                            SessionId = session.Id,
                            SessionTimestampUtc = utcNow,

                            ProfileName = profileName,
                            ScopeType = scopeType.ToString(),
                            ScopeDescription = scopeDescription,

                            DocumentFile = export.IncludeDocumentFile ? identity.DocumentFile : null,
                            DocumentFileKey = export.IncludeDocumentFile ? identity.DocumentFileKey : null,
                            SourceFile = export.IncludeSourceFile ? identity.SourceFile : null,
                            SourceFileKey = export.IncludeSourceFile ? identity.SourceFileKey : null,

                            PrimaryKey = identity.PrimaryKey,
                            PrimaryKeyType = identity.PrimaryKeyType,

                            NavisInstanceGuid = export.IncludeNavisworksInstanceGuid ? identity.NavisInstanceGuid : null,
                            RevitUniqueId = export.IncludeRevitIds ? identity.RevitUniqueId : null,
                            RevitElementId = export.IncludeRevitIds ? identity.RevitElementId : null,
                            IfcGlobalId = export.IncludeIfcIds ? identity.IfcGlobalId : null,
                            DwgHandle = export.IncludeDwgIds ? identity.DwgHandle : null,

                            Ids = identity.Ids,
                            ItemPath = itemPath,

                            PropertyCount = docProperties.Count,
                            Properties = docProperties
                        };

                        if (itemWriter != null)
                        {
                            itemWriter.Write(record);
                        }
                        else
                        {
                            pendingItemDocs?.Add(record);
                        }
                    }
                }
            }
            finally
            {
                if (rawWriter != null)
                {
                    rawWriter.Dispose();
                }
                if (itemWriter != null)
                {
                    itemWriter.Dispose();
                }
            }

            if (exportEnabled && !exportStreaming && (pendingRawRows != null || pendingItemDocs != null))
            {
                if (pendingRawRows != null && rawExportPath != null)
                {
                    using var exportWriter = new JsonlStreamWriter(rawExportPath, export.Gzip);
                    foreach (var record in pendingRawRows)
                    {
                        exportWriter.Write(record);
                    }
                    session.JsonlLinesWritten += exportWriter.LinesWritten;
                }

                if (pendingItemDocs != null && itemExportPath != null)
                {
                    using var exportWriter = new JsonlStreamWriter(itemExportPath, export.Gzip);
                    foreach (var record in pendingItemDocs)
                    {
                        exportWriter.Write(record);
                    }
                    session.JsonlLinesWritten += exportWriter.LinesWritten;
                }
            }
            else if (exportStreaming)
            {
                session.JsonlLinesWritten = (rawWriter?.LinesWritten ?? 0) + (itemWriter?.LinesWritten ?? 0);
            }

            if (exportEnabled && export.ExportSource == DataScraperExportSource.PropertySummaries)
            {
                var summaryLines = ExportPropertySummaries(session, export, summaryExportPath);
                session.JsonlLinesWritten += summaryLines;
            }

            session.Properties = propertyMap.Values
                .Select(a => new ScrapedProperty
                {
                    Category = a.Category,
                    Name = a.Name,
                    DataType = a.DataType,
                    ItemCount = a.ItemCount,
                    DistinctValueCount = a.Distinct.Count,
                    SampleValues = a.Sample
                })
                .OrderBy(p => p.Category)
                .ThenBy(p => p.Name)
                .ToList();

            if (exportEnabled && session.ItemsScanned > 0)
            {
                var n = session.ItemsScanned;
                string Percent(int count) => $"{Math.Round(count * 100.0 / n):0}%";

                MicroEngActions.Log(
                    $"Identity capture: sourceFile={Percent(sourceFileHit)}, revitUniqueId={Percent(revitUniqueIdHit)}, " +
                    $"ifcGlobalId={Percent(ifcGlobalIdHit)}, dwgHandle={Percent(dwgHandleHit)} (n={n})");
            }

            DataScraperCache.AddSession(session);
            return session;
        }

        public long ExportFromSession(ScrapeSession session, DataScraperJsonlExportSettings export)
        {
            if (session == null)
            {
                throw new ArgumentNullException(nameof(session));
            }

            if (export == null)
            {
                throw new ArgumentNullException(nameof(export));
            }

            var exportRaw = export.ExportRawRows;
            var exportItems = export.ExportItemDocuments;
            if (!exportRaw && !exportItems)
            {
                if (export.ExportSource != DataScraperExportSource.PropertySummaries)
                {
                    return 0;
                }
            }

            var utcNow = DateTime.UtcNow;
            var rawPath = exportRaw
                ? BuildExportPath(export.OutputPath, session.ProfileName ?? "Profile", utcNow, export.Gzip, exportItems ? "_raw" : string.Empty)
                : null;
            var itemPath = exportItems
                ? BuildExportPath(export.OutputPath, session.ProfileName ?? "Profile", utcNow, export.Gzip, exportRaw ? "_items" : string.Empty)
                : null;

            var timestampUtc = session.Timestamp.Kind == DateTimeKind.Utc
                ? session.Timestamp
                : session.Timestamp.ToUniversalTime();

            JsonlStreamWriter rawWriter = null;
            JsonlStreamWriter itemWriter = null;

            if (rawPath != null)
            {
                rawWriter = new JsonlStreamWriter(rawPath, export.Gzip);
            }

            if (itemPath != null)
            {
                itemWriter = new JsonlStreamWriter(itemPath, export.Gzip);
            }

            if (export.ExportSource == DataScraperExportSource.PropertySummaries)
            {
                var summaryPath = BuildExportPath(export.OutputPath, session.ProfileName ?? "Profile", utcNow, export.Gzip, "_summary");
                var lines = ExportPropertySummaries(session, export, summaryPath);
                session.JsonlExportEnabled = true;
                session.JsonlExportPath = summaryPath;
                session.JsonlExportMode = "PropertySummaries";
                session.JsonlExportGzip = export.Gzip;
                session.JsonlPrimaryKeyMode = export.PrimaryKeyMode.ToString();
                session.JsonlLinesWritten = lines;
                session.JsonlStreamedDuringScrape = false;
                return lines;
            }

            var groupedDocs = exportItems
                ? BuildItemDocumentsFromRaw(session.RawEntries ?? new List<RawEntry>())
                : new Dictionary<string, JsonlItemDocRecord>();

            foreach (var entry in session.RawEntries ?? Enumerable.Empty<RawEntry>())
            {
                if (exportRaw && rawWriter != null)
                {
                    var parsed = ParseValue(entry.Value);

                    var record = new JsonlRawRowRecord
                    {
                        SessionId = session.Id,
                        SessionTimestampUtc = timestampUtc,

                        ProfileName = string.IsNullOrWhiteSpace(entry.Profile) ? session.ProfileName : entry.Profile,
                        ScopeType = session.ScopeType,
                        ScopeDescription = session.ScopeDescription,

                        DocumentFile = export.IncludeDocumentFile ? session.DocumentFile : null,
                        DocumentFileKey = export.IncludeDocumentFile ? session.DocumentFileKey : null,
                        SourceFile = null,
                        SourceFileKey = null,

                        PrimaryKey = GetPrimaryKey(entry.ItemKey, entry.ItemPath, export.PrimaryKeyMode),
                        PrimaryKeyType = export.PrimaryKeyMode.ToString(),

                        NavisInstanceGuid = export.IncludeNavisworksInstanceGuid ? entry.ItemKey : null,
                        RevitUniqueId = null,
                        RevitElementId = null,
                        IfcGlobalId = null,
                        DwgHandle = null,

                        Ids = null,
                        ItemPath = entry.ItemPath,

                        Category = entry.Category,
                        Property = entry.Name,
                        DataType = entry.DataType,
                        Value = entry.Value,
                        ValueNorm = parsed.ValueNorm,
                        ValueNum = parsed.ValueNum,
                        ValueBool = parsed.ValueBool,
                        ValueDateUtc = parsed.ValueDateUtc
                    };

                    rawWriter.Write(record);
                }
            }

            if (exportItems && itemWriter != null)
            {
                foreach (var doc in groupedDocs.Values)
                {
                    doc.SessionId = session.Id;
                    doc.SessionTimestampUtc = timestampUtc;
                    doc.ProfileName = session.ProfileName;
                    doc.ScopeType = session.ScopeType;
                    doc.ScopeDescription = session.ScopeDescription;
                    doc.DocumentFile = export.IncludeDocumentFile ? session.DocumentFile : null;
                    doc.DocumentFileKey = export.IncludeDocumentFile ? session.DocumentFileKey : null;
                    doc.SourceFile = null;
                    doc.SourceFileKey = null;
                    doc.PrimaryKey = GetPrimaryKey(doc.NavisInstanceGuid, doc.ItemPath, export.PrimaryKeyMode);
                    doc.PrimaryKeyType = export.PrimaryKeyMode.ToString();
                    doc.NavisInstanceGuid = export.IncludeNavisworksInstanceGuid ? doc.NavisInstanceGuid : null;

                    itemWriter.Write(doc);
                }
            }

            var linesWritten = (rawWriter?.LinesWritten ?? 0) + (itemWriter?.LinesWritten ?? 0);

            rawWriter?.Dispose();
            itemWriter?.Dispose();

            session.JsonlExportEnabled = true;
            session.JsonlExportPath = exportRaw && exportItems
                ? $"{rawPath} | {itemPath}"
                : rawPath ?? itemPath;
            session.JsonlExportMode = exportRaw && exportItems
                ? "RawRows+ItemDocuments"
                : exportItems
                    ? DataScraperJsonlExportMode.ItemDocuments.ToString()
                    : DataScraperJsonlExportMode.RawRows.ToString();
            session.JsonlExportGzip = export.Gzip;
            session.JsonlPrimaryKeyMode = export.PrimaryKeyMode.ToString();
            session.JsonlLinesWritten = linesWritten;
            session.JsonlStreamedDuringScrape = false;

            return linesWritten;
        }

        private static string GetPrimaryKey(string itemKey, string itemPath, DataScraperPrimaryKeyMode mode)
        {
            switch (mode)
            {
                case DataScraperPrimaryKeyMode.ItemPath:
                    return itemPath ?? string.Empty;
                case DataScraperPrimaryKeyMode.InstanceGuid:
                case DataScraperPrimaryKeyMode.BestExternalId:
                case DataScraperPrimaryKeyMode.SourceFilePlusBestExternalId:
                default:
                    return itemKey ?? string.Empty;
            }
        }

        private static Dictionary<string, JsonlItemDocRecord> BuildItemDocumentsFromRaw(List<RawEntry> rawEntries)
        {
            var docs = new Dictionary<string, JsonlItemDocRecord>(StringComparer.OrdinalIgnoreCase);
            if (rawEntries == null)
            {
                return docs;
            }

            foreach (var entry in rawEntries)
            {
                if (entry == null)
                {
                    continue;
                }

                var key = entry.ItemKey ?? entry.ItemPath ?? string.Empty;
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                if (!docs.TryGetValue(key, out var doc))
                {
                    doc = new JsonlItemDocRecord
                    {
                        NavisInstanceGuid = entry.ItemKey,
                        ItemPath = entry.ItemPath,
                        Properties = new List<JsonlItemDocProperty>()
                    };
                    docs[key] = doc;
                }

                var propKey = $"{entry.Category}\u001F{entry.Name}\u001F{entry.DataType}";
                var prop = doc.Properties.FirstOrDefault(p => string.Equals($"{p.Category}\u001F{p.Property}\u001F{p.DataType}", propKey, StringComparison.OrdinalIgnoreCase));
                if (prop == null)
                {
                    prop = new JsonlItemDocProperty
                    {
                        Category = entry.Category,
                        Property = entry.Name,
                        DataType = entry.DataType,
                        Values = new List<string>()
                    };
                    doc.Properties.Add(prop);
                }

                if (!string.IsNullOrWhiteSpace(entry.Value) && !prop.Values.Contains(entry.Value))
                {
                    prop.Values.Add(entry.Value);
                }
            }

            foreach (var doc in docs.Values)
            {
                doc.PropertyCount = doc.Properties?.Count ?? 0;
            }

            return docs;
        }

        private static long ExportPropertySummaries(ScrapeSession session, DataScraperJsonlExportSettings export, string outputPath)
        {
            if (session == null || string.IsNullOrWhiteSpace(outputPath))
            {
                return 0;
            }

            using var writer = new JsonlStreamWriter(outputPath, export.Gzip);
            var timestampUtc = session.Timestamp.Kind == DateTimeKind.Utc
                ? session.Timestamp
                : session.Timestamp.ToUniversalTime();

            foreach (var prop in session.Properties ?? Enumerable.Empty<ScrapedProperty>())
            {
                var record = new JsonlPropertySummaryRecord
                {
                    SessionId = session.Id,
                    SessionTimestampUtc = timestampUtc,
                    ProfileName = session.ProfileName,
                    ScopeType = session.ScopeType,
                    ScopeDescription = session.ScopeDescription,
                    Category = prop.Category,
                    Property = prop.Name,
                    DataType = prop.DataType,
                    Items = prop.ItemCount,
                    Distinct = prop.DistinctValueCount,
                    Samples = prop.SampleValues
                };
                writer.Write(record);
            }

            return writer.LinesWritten;
        }

        private static string BuildExportPath(string outputPath, string profileName, DateTime utcNow, bool gzip, string suffix)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                var safeProfile = string.IsNullOrWhiteSpace(profileName) ? "Profile" : profileName.Trim();
                if (!string.IsNullOrWhiteSpace(suffix))
                {
                    safeProfile += suffix;
                }
                return GetDefaultJsonlPath(safeProfile, gzip, utcNow);
            }

            var basePath = EnsureJsonlExtension(outputPath, gzip);
            if (string.IsNullOrWhiteSpace(suffix))
            {
                return basePath;
            }

            var ext = gzip ? ".jsonl.gz" : ".jsonl";
            if (basePath.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
            {
                return basePath.Substring(0, basePath.Length - ext.Length) + suffix + ext;
            }

            return basePath + suffix;
        }

        public IEnumerable<ModelItem> ResolveScope(ScrapeScopeType scopeType, string selectionSetName, string searchSetName, out string description)
        {
            description = string.Empty;
            var doc = NavisApp.ActiveDocument;
            if (doc == null) return Enumerable.Empty<ModelItem>();

            switch (scopeType)
            {
                case ScrapeScopeType.SingleItem:
                    var single = doc.CurrentSelection?.SelectedItems?.FirstOrDefault();
                    if (single != null)
                    {
                        description = $"Single Item: {single.DisplayName}";
                        return new[] { single };
                    }
                    return Enumerable.Empty<ModelItem>();
                case ScrapeScopeType.CurrentSelection:
                    var sel = doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                    description = $"Current Selection ({sel.Count()} items)";
                    return sel;
                case ScrapeScopeType.SelectionSet:
                    var selectionItems = ResolveSelectionSetInternal(selectionSetName).ToList();
                    description = string.IsNullOrWhiteSpace(selectionSetName)
                        ? $"Selection Set ({selectionItems.Count} items)"
                        : $"Selection Set: {selectionSetName} ({selectionItems.Count} items)";
                    return selectionItems;
                case ScrapeScopeType.SearchSet:
                    var searchItems = ResolveSelectionSetInternal(searchSetName).ToList();
                    description = string.IsNullOrWhiteSpace(searchSetName)
                        ? $"Search Set ({searchItems.Count} items)"
                        : $"Search Set: {searchSetName} ({searchItems.Count} items)";
                    return searchItems;
                case ScrapeScopeType.EntireModel:
                    description = "Entire Model";
                    return Traverse(doc.Models.RootItems);
                default:
                    return Enumerable.Empty<ModelItem>();
            }
        }

        private static IEnumerable<ModelItem> ResolveSelectionSetInternal(string setName)
        {
            if (string.IsNullOrWhiteSpace(setName))
            {
                return Enumerable.Empty<ModelItem>();
            }

            try
            {
                var doc = NavisApp.ActiveDocument;
                var root = doc?.SelectionSets?.RootItem;
                var selectionSet = FindSelectionSetByName(root, setName);
                if (selectionSet != null)
                {
                    var items = selectionSet.GetSelectedItems(doc);
                    return items?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                }
            }
            catch
            {
                // ignore managed API selection set failures
            }

            try
            {
                var state = ComBridge.State;
                var sets = state.SelectionSetsEx();
                var selectionSet = FindSelectionSetByName(sets, setName);
                if (selectionSet == null)
                {
                    return Enumerable.Empty<ModelItem>();
                }

                var items = ComBridge.ToModelItemCollection(selectionSet.selection);
                return items?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
            }
            catch
            {
                return Enumerable.Empty<ModelItem>();
            }
        }

        private static SelectionSet FindSelectionSetByName(FolderItem folder, string name)
        {
            if (folder == null || string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            var children = folder.Children;
            if (children == null)
            {
                return null;
            }

            foreach (SavedItem item in children)
            {
                if (item is SelectionSet set)
                {
                    if (string.Equals(set.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return set;
                    }
                }
                else if (item is FolderItem subFolder)
                {
                    var found = FindSelectionSetByName(subFolder, name);
                    if (found != null)
                    {
                        return found;
                    }
                }
            }

            return null;
        }

        private static ComApi.InwOpSelectionSet FindSelectionSetByName(ComApi.InwSelectionSetExColl collection, string name)
        {
            if (collection == null || string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            for (int i = 1; i <= collection.Count; i++)
            {
                var item = collection[i];
                if (item is ComApi.InwOpSelectionSet set)
                {
                    if (string.Equals(set.name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return set;
                    }
                }
                else if (item is ComApi.InwSelectionSetFolder folder)
                {
                    var found = FindSelectionSetByName(folder.SelectionSets(), name);
                    if (found != null)
                    {
                        return found;
                    }
                }
            }

            return null;
        }

        private static IEnumerable<ModelItem> Traverse(IEnumerable<ModelItem> items)
        {
            foreach (ModelItem item in items)
            {
                yield return item;
                if (item.Children != null && item.Children.Any())
                {
                    foreach (var child in Traverse(item.Children))
                        yield return child;
                }
            }
        }

        private static string ItemToPath(ModelItem item)
        {
            if (item == null)
            {
                return string.Empty;
            }

            try
            {
                var names = new List<string>();
                var current = item;
                var guard = 0;
                while (current != null && guard++ < 128)
                {
                    var name = current.TryGetDisplayName();
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        names.Add(name);
                    }

                    current = current.Parent;
                }

                names.Reverse();
                if (names.Count == 0)
                {
                    return item.InstanceGuid.ToString("D");
                }

                return string.Join(" / ", names);
            }
            catch
            {
                return item.TryGetDisplayName();
            }
        }

        private static List<string> GetPropertyValueStrings(DataProperty prop)
        {
            var values = new List<string>();
            if (prop?.Value == null)
            {
                return values;
            }

            try
            {
                if (prop.Value.IsDisplayString)
                {
                    var display = prop.Value.ToDisplayString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(display))
                    {
                        values.Add(display);
                    }

                    return values;
                }
            }
            catch
            {
                // ignore display string failures
            }

            try
            {
                var text = prop.Value.ToString() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    values.Add(text);
                }
            }
            catch
            {
                // ignore
            }

            return values;
        }

        private static void TouchProperty(
            Dictionary<string, PropertyAccumulator> map,
            string category,
            string name,
            string dataType,
            bool firstThisItem,
            List<string> values)
        {
            var key = $"{category}\u001F{name}\u001F{dataType}";
            if (!map.TryGetValue(key, out var acc))
            {
                acc = new PropertyAccumulator
                {
                    Category = category,
                    Name = name,
                    DataType = dataType
                };
                map[key] = acc;
            }

            if (firstThisItem)
            {
                acc.ItemCount++;
            }

            foreach (var v in values)
            {
                if (acc.Distinct.Count < 5000)
                {
                    acc.Distinct.Add(v);
                }

                if (acc.Sample.Count < 12 && !acc.Sample.Any(s => string.Equals(s, v, StringComparison.OrdinalIgnoreCase)))
                {
                    acc.Sample.Add(v);
                }
            }
        }

        private static ParsedValue ParseValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return default;
            }

            var trimmed = value.Trim();
            if (trimmed.Length == 0)
            {
                return default;
            }

            var norm = trimmed.ToLowerInvariant();
            double? number = null;
            bool? boolean = null;
            DateTime? dateUtc = null;

            if (double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var num))
            {
                number = num;
            }

            if (TryParseBool(trimmed, out var boolVal))
            {
                boolean = boolVal;
            }

            if (DateTime.TryParse(trimmed, CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeLocal | DateTimeStyles.AdjustToUniversal, out var dt))
            {
                dateUtc = dt;
            }

            return new ParsedValue(norm, number, boolean, dateUtc);
        }

        private static IdentityResult BuildIdentity(
            ModelItem item,
            string itemPath,
            string documentFile,
            IdentityProbe probe,
            DataScraperJsonlExportSettings export)
        {
            var navisGuid = item?.InstanceGuid.ToString("D") ?? string.Empty;
            var sourceFile = probe?.GetValue("sourceFile") ?? string.Empty;
            var revitUniqueId = probe?.GetValue("revitUniqueId") ?? string.Empty;
            var revitElementId = probe?.GetValue("revitElementId") ?? string.Empty;
            var ifcGlobalId = probe?.GetValue("ifcGlobalId") ?? string.Empty;
            var dwgHandle = probe?.GetValue("dwgHandle") ?? string.Empty;

            navisGuid = NormalizeGuidLike(navisGuid, export.NormalizeIds);
            revitUniqueId = export.NormalizeIds ? NormalizeRevitUniqueId(revitUniqueId) : revitUniqueId?.Trim() ?? string.Empty;
            revitElementId = export.NormalizeIds ? NormalizeGuidLike(revitElementId) : revitElementId?.Trim() ?? string.Empty;
            ifcGlobalId = ifcGlobalId?.Trim() ?? string.Empty;
            dwgHandle = dwgHandle?.Trim() ?? string.Empty;

            var documentFileClean = documentFile?.Trim() ?? string.Empty;
            var documentKey = BuildPathKey(documentFileClean, export.SourceFileKeyMode);
            var sourceKey = BuildPathKey(sourceFile, export.SourceFileKeyMode);

            var bestExternal = FirstNonEmpty(revitUniqueId, ifcGlobalId, revitElementId, dwgHandle, navisGuid);
            if (string.IsNullOrWhiteSpace(bestExternal))
            {
                bestExternal = navisGuid;
            }

            string primaryKey;
            var primaryKeyType = export.PrimaryKeyMode.ToString();
            switch (export.PrimaryKeyMode)
            {
                case DataScraperPrimaryKeyMode.InstanceGuid:
                    primaryKey = navisGuid;
                    break;
                case DataScraperPrimaryKeyMode.BestExternalId:
                    primaryKey = bestExternal;
                    break;
                case DataScraperPrimaryKeyMode.SourceFilePlusBestExternalId:
                    primaryKey = !string.IsNullOrWhiteSpace(sourceKey)
                        ? $"{sourceKey}|{bestExternal}"
                        : bestExternal;
                    break;
                case DataScraperPrimaryKeyMode.ItemPath:
                    primaryKey = itemPath;
                    break;
                default:
                    primaryKey = navisGuid;
                    break;
            }

            if (string.IsNullOrWhiteSpace(primaryKey))
            {
                primaryKey = navisGuid;
            }
            if (string.IsNullOrWhiteSpace(primaryKey))
            {
                primaryKey = itemPath ?? string.Empty;
            }

            var ids = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (export.IncludeNavisworksInstanceGuid && !string.IsNullOrWhiteSpace(navisGuid))
            {
                ids["navis.instanceGuid"] = navisGuid;
            }
            if (export.IncludeRevitIds)
            {
                if (!string.IsNullOrWhiteSpace(revitUniqueId))
                {
                    ids["revit.uniqueId"] = revitUniqueId;
                }
                if (!string.IsNullOrWhiteSpace(revitElementId))
                {
                    ids["revit.elementId"] = revitElementId;
                }
            }
            if (export.IncludeIfcIds && !string.IsNullOrWhiteSpace(ifcGlobalId))
            {
                ids["ifc.globalId"] = ifcGlobalId;
            }
            if (export.IncludeDwgIds && !string.IsNullOrWhiteSpace(dwgHandle))
            {
                ids["dwg.handle"] = dwgHandle;
            }

            return new IdentityResult
            {
                DocumentFile = string.IsNullOrWhiteSpace(documentFileClean) ? null : documentFileClean,
                DocumentFileKey = string.IsNullOrWhiteSpace(documentKey) ? null : documentKey,
                SourceFile = string.IsNullOrWhiteSpace(sourceFile) ? null : sourceFile,
                SourceFileKey = string.IsNullOrWhiteSpace(sourceKey) ? null : sourceKey,
                PrimaryKey = primaryKey,
                PrimaryKeyType = primaryKeyType,
                NavisInstanceGuid = navisGuid,
                RevitUniqueId = revitUniqueId,
                RevitElementId = revitElementId,
                IfcGlobalId = ifcGlobalId,
                DwgHandle = dwgHandle,
                Ids = ids.Count > 0 ? ids : null
            };
        }

        private static string BuildPathKey(string path, DataScraperPathKeyMode mode)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            var trimmed = path.Trim();
            if (mode == DataScraperPathKeyMode.FullPath)
            {
                return trimmed;
            }

            return Path.GetFileName(trimmed) ?? trimmed;
        }

        private static string EnsureJsonlExtension(string path, bool gzip)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return path;
            }

            var p = path.Trim();

            if (gzip)
            {
                if (p.EndsWith(".jsonl.gz", StringComparison.OrdinalIgnoreCase)) return p;
                if (p.EndsWith(".gz", StringComparison.OrdinalIgnoreCase)) return p;
                if (p.EndsWith(".jsonl", StringComparison.OrdinalIgnoreCase)) return p + ".gz";
                return p + ".jsonl.gz";
            }

            if (p.EndsWith(".jsonl", StringComparison.OrdinalIgnoreCase)) return p;
            if (p.EndsWith(".jsonl.gz", StringComparison.OrdinalIgnoreCase))
            {
                return p.Substring(0, p.Length - 3);
            }

            return p + ".jsonl";
        }

        private static string GetDefaultJsonlPath(string profileName, bool gzip, DateTime utcNow)
        {
            var safeProfile = MakeSafeFileName(profileName);
            var ts = utcNow.ToString("yyyyMMdd_HHmmss");
            var ext = gzip ? ".jsonl.gz" : ".jsonl";

            var folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MicroEng",
                "DataScraper",
                "Exports");

            Directory.CreateDirectory(folder);

            return Path.Combine(folder, $"DataScraper_{safeProfile}_{ts}{ext}");
        }

        private static string MakeSafeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "Profile";
            }

            var safe = name;
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                safe = safe.Replace(c, '_');
            }

            return safe.Trim();
        }

        private static string NormalizeGuidLike(string value, bool normalize = true)
        {
            if (!normalize)
            {
                return value?.Trim() ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var trimmed = value.Trim();
            if (Guid.TryParse(trimmed, out var guid))
            {
                return guid.ToString("D");
            }

            return trimmed;
        }

        private static string NormalizeRevitUniqueId(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var trimmed = value.Trim();
            if (trimmed.Length < 36)
            {
                return trimmed;
            }

            var guidPart = trimmed.Substring(0, 36);
            if (Guid.TryParse(guidPart, out var guid))
            {
                var suffix = trimmed.Length > 36 ? trimmed.Substring(36) : string.Empty;
                return guid.ToString("D") + suffix;
            }

            return trimmed;
        }

        private static string FirstNonEmpty(params string[] values)
        {
            foreach (var v in values)
            {
                if (!string.IsNullOrWhiteSpace(v))
                {
                    return v;
                }
            }

            return string.Empty;
        }

        private static bool TryParseBool(string value, out bool result)
        {
            if (bool.TryParse(value, out result))
            {
                return true;
            }

            if (string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase))
            {
                result = true;
                return true;
            }

            if (string.Equals(value, "no", StringComparison.OrdinalIgnoreCase))
            {
                result = false;
                return true;
            }

            if (string.Equals(value, "1", StringComparison.OrdinalIgnoreCase))
            {
                result = true;
                return true;
            }

            if (string.Equals(value, "0", StringComparison.OrdinalIgnoreCase))
            {
                result = false;
                return true;
            }

            return false;
        }

    }

    internal sealed class JsonlItemDocProperty
    {
        public string Category { get; set; }
        public string Property { get; set; }
        public string DataType { get; set; }
        public List<string> Values { get; set; }
    }

    internal sealed class JsonlRawRowRecord
    {
        public int SchemaVersion { get; set; } = 1;
        public string RecordType { get; set; } = "rawRow";

        public Guid SessionId { get; set; }
        public DateTime SessionTimestampUtc { get; set; }

        public string ProfileName { get; set; }
        public string ScopeType { get; set; }
        public string ScopeDescription { get; set; }

        public string DocumentFile { get; set; }
        public string DocumentFileKey { get; set; }
        public string SourceFile { get; set; }
        public string SourceFileKey { get; set; }

        public string PrimaryKey { get; set; }
        public string PrimaryKeyType { get; set; }

        public string NavisInstanceGuid { get; set; }
        public string RevitUniqueId { get; set; }
        public string RevitElementId { get; set; }
        public string IfcGlobalId { get; set; }
        public string DwgHandle { get; set; }

        public Dictionary<string, string> Ids { get; set; }
        public string ItemPath { get; set; }

        public string Category { get; set; }
        public string Property { get; set; }
        public string DataType { get; set; }

        public string Value { get; set; }
        public string ValueNorm { get; set; }
        public double? ValueNum { get; set; }
        public bool? ValueBool { get; set; }
        public DateTime? ValueDateUtc { get; set; }
    }

    internal sealed class JsonlItemDocRecord
    {
        public int SchemaVersion { get; set; } = 1;
        public string RecordType { get; set; } = "itemDoc";

        public Guid SessionId { get; set; }
        public DateTime SessionTimestampUtc { get; set; }

        public string ProfileName { get; set; }
        public string ScopeType { get; set; }
        public string ScopeDescription { get; set; }

        public string DocumentFile { get; set; }
        public string DocumentFileKey { get; set; }
        public string SourceFile { get; set; }
        public string SourceFileKey { get; set; }

        public string PrimaryKey { get; set; }
        public string PrimaryKeyType { get; set; }

        public string NavisInstanceGuid { get; set; }
        public string RevitUniqueId { get; set; }
        public string RevitElementId { get; set; }
        public string IfcGlobalId { get; set; }
        public string DwgHandle { get; set; }

        public Dictionary<string, string> Ids { get; set; }
        public string ItemPath { get; set; }

        public int PropertyCount { get; set; }
        public List<JsonlItemDocProperty> Properties { get; set; }
    }

    internal sealed class JsonlPropertySummaryRecord
    {
        public int SchemaVersion { get; set; } = 1;
        public string RecordType { get; set; } = "propertySummary";

        public Guid SessionId { get; set; }
        public DateTime SessionTimestampUtc { get; set; }

        public string ProfileName { get; set; }
        public string ScopeType { get; set; }
        public string ScopeDescription { get; set; }

        public string Category { get; set; }
        public string Property { get; set; }
        public string DataType { get; set; }
        public int Items { get; set; }
        public int Distinct { get; set; }
        public List<string> Samples { get; set; }
    }

    internal sealed class JsonlStreamWriter : IDisposable
    {
        private readonly StreamWriter _writer;
        private readonly JsonSerializer _ser;
        public long LinesWritten { get; private set; }

        public JsonlStreamWriter(string path, bool gzip)
        {
            Stream stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read);
            if (gzip)
            {
                stream = new GZipStream(stream, CompressionLevel.Optimal, leaveOpen: false);
            }

            _writer = new StreamWriter(stream, new UTF8Encoding(false), 1 << 16, leaveOpen: false);
            _ser = JsonSerializer.Create(new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                Formatting = Formatting.None,
                ContractResolver = new CamelCasePropertyNamesContractResolver()
            });
        }

        public void Write(object record)
        {
            _ser.Serialize(_writer, record);
            _writer.WriteLine();
            LinesWritten++;

            if ((LinesWritten % 4096) == 0)
            {
                _writer.Flush();
            }
        }

        public void Dispose()
        {
            try { _writer.Flush(); } catch { }
            _writer.Dispose();
        }
    }

    internal sealed class IdentityResult
    {
        public string DocumentFile { get; set; }
        public string DocumentFileKey { get; set; }
        public string SourceFile { get; set; }
        public string SourceFileKey { get; set; }
        public string PrimaryKey { get; set; }
        public string PrimaryKeyType { get; set; }
        public string NavisInstanceGuid { get; set; }
        public string RevitUniqueId { get; set; }
        public string RevitElementId { get; set; }
        public string IfcGlobalId { get; set; }
        public string DwgHandle { get; set; }
        public Dictionary<string, string> Ids { get; set; }
    }

    internal sealed class ItemRawRow
    {
        public string Category { get; set; }
        public string Property { get; set; }
        public string DataType { get; set; }
        public string Value { get; set; }
    }

    internal sealed class PropertyAccumulator
    {
        public string Category { get; set; }
        public string Name { get; set; }
        public string DataType { get; set; }
        public int ItemCount { get; set; }
        public HashSet<string> Distinct { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<string> Sample { get; } = new();
    }

    internal readonly struct ParsedValue
    {
        public ParsedValue(string valueNorm, double? valueNum, bool? valueBool, DateTime? valueDateUtc)
        {
            ValueNorm = valueNorm;
            ValueNum = valueNum;
            ValueBool = valueBool;
            ValueDateUtc = valueDateUtc;
        }

        public string ValueNorm { get; }
        public double? ValueNum { get; }
        public bool? ValueBool { get; }
        public DateTime? ValueDateUtc { get; }
    }

    internal sealed class IdentityProbe
    {
        private readonly Dictionary<string, ProbeHit> _bestHits = new(StringComparer.OrdinalIgnoreCase);
        private static readonly List<ProbeRule> Rules = new()
        {
            new ProbeRule("sourceFile", "item", "sourcefilename", 100),
            new ProbeRule("sourceFile", "item", "lcoapartitionsourcefilename", 95),
            new ProbeRule("sourceFile", "item", "sourcefile", 90),
            new ProbeRule("sourceFile", "item", "filename", 40),

            new ProbeRule("revitUniqueId", "elementproperties", "uniqueid", 100),
            new ProbeRule("revitElementId", "revitelementid", "value", 100),
            new ProbeRule("revitElementId", "elementproperties", "elementid", 70),

            new ProbeRule("ifcGlobalId", "ifc", "globalid", 100),
            new ProbeRule("ifcGlobalId", "ifc", "ifcguid", 90),

            new ProbeRule("dwgHandle", "autocadentityhandle", "value", 100),
            new ProbeRule("dwgHandle", "entityhandle", "value", 90),
            new ProbeRule("dwgHandle", "autocad", "handle", 80)
        };

        public void Observe(string categoryDisplay, string categoryInternal, string propertyDisplay, string propertyInternal, List<string> values)
        {
            if (values == null || values.Count == 0)
            {
                return;
            }

            var catKey = NormalizeKey(categoryDisplay);
            var catInternalKey = NormalizeKey(categoryInternal);
            var propKey = NormalizeKey(propertyDisplay);
            var propInternalKey = NormalizeKey(propertyInternal);

            foreach (var rule in Rules)
            {
                if (!string.Equals(rule.CategoryKey, catKey, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(rule.CategoryKey, catInternalKey, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.Equals(rule.PropertyKey, propKey, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(rule.PropertyKey, propInternalKey, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                foreach (var value in values)
                {
                    var trimmed = value?.Trim() ?? string.Empty;
                    if (trimmed.Length == 0)
                    {
                        continue;
                    }

                    if (!IsValid(rule.FieldKey, trimmed))
                    {
                        continue;
                    }

                    if (_bestHits.TryGetValue(rule.FieldKey, out var hit) && hit.Score >= rule.Weight)
                    {
                        continue;
                    }

                    _bestHits[rule.FieldKey] = new ProbeHit(rule.FieldKey, rule.Weight, trimmed);
                }
            }
        }

        public string GetValue(string fieldKey)
        {
            return _bestHits.TryGetValue(fieldKey, out var hit) ? hit.Value : string.Empty;
        }

        private static bool IsValid(string fieldKey, string value)
        {
            switch (fieldKey)
            {
                case "sourceFile":
                    return IsLikelySourceFile(value);
                case "revitUniqueId":
                    return LooksLikeRevitUniqueId(value);
                case "revitElementId":
                    return LooksLikeNumeric(value);
                case "ifcGlobalId":
                    return LooksLikeIfcGlobalId(value);
                case "dwgHandle":
                    return LooksLikeHex(value);
                default:
                    return !string.IsNullOrWhiteSpace(value);
            }
        }

        private static bool IsLikelySourceFile(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var v = value.Trim();
            if (v.IndexOfAny(new[] { '\\', '/' }) >= 0)
            {
                return true;
            }

            var lower = v.ToLowerInvariant();
            var exts = new[] { ".nwc", ".nwd", ".nwf", ".rvt", ".ifc", ".dwg", ".dxf" };
            return exts.Any(ext => lower.EndsWith(ext, StringComparison.OrdinalIgnoreCase) || lower.Contains(ext));
        }

        private static bool LooksLikeRevitUniqueId(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var trimmed = value.Trim();
            if (trimmed.Length < 36)
            {
                return false;
            }

            var guidPart = trimmed.Substring(0, 36);
            return Guid.TryParse(guidPart, out _);
        }

        private static bool LooksLikeIfcGlobalId(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var trimmed = value.Trim();
            if (trimmed.Length < 18 || trimmed.Length > 30)
            {
                return false;
            }

            foreach (var c in trimmed)
            {
                if (!(char.IsLetterOrDigit(c) || c == '_' || c == '$'))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool LooksLikeHex(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var trimmed = value.Trim();
            if (trimmed.Length < 4 || trimmed.Length > 16)
            {
                return false;
            }

            foreach (var c in trimmed)
            {
                var isHex = (c >= '0' && c <= '9')
                            || (c >= 'a' && c <= 'f')
                            || (c >= 'A' && c <= 'F');
                if (!isHex)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool LooksLikeNumeric(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var trimmed = value.Trim();
            var start = 0;
            if (trimmed.StartsWith("-", StringComparison.Ordinal) || trimmed.StartsWith("+", StringComparison.Ordinal))
            {
                start = 1;
            }

            if (trimmed.Length <= start)
            {
                return false;
            }

            for (var i = start; i < trimmed.Length; i++)
            {
                if (!char.IsDigit(trimmed[i]))
                {
                    return false;
                }
            }

            return true;
        }

        private static string NormalizeKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(value.Length);
            foreach (var c in value)
            {
                if (char.IsLetterOrDigit(c))
                {
                    sb.Append(char.ToLowerInvariant(c));
                }
            }

            return sb.ToString();
        }
    }

    internal readonly struct ProbeRule
    {
        public ProbeRule(string fieldKey, string categoryKey, string propertyKey, int weight)
        {
            FieldKey = fieldKey;
            CategoryKey = categoryKey;
            PropertyKey = propertyKey;
            Weight = weight;
        }

        public string FieldKey { get; }
        public string CategoryKey { get; }
        public string PropertyKey { get; }
        public int Weight { get; }
    }

    internal readonly struct ProbeHit
    {
        public ProbeHit(string fieldKey, int score, string value)
        {
            FieldKey = fieldKey;
            Score = score;
            Value = value;
        }

        public string FieldKey { get; }
        public int Score { get; }
        public string Value { get; }
    }

    internal static class ModelItemExtensions
    {
        public static string TryGetDisplayName(this ModelItem item)
        {
            try
            {
                return item?.DisplayName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
