using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading;

namespace MicroEng.Navisworks
{
    [DataContract]
    internal class ScrapedProperty
    {
        [DataMember(Order = 1)]
        public string Category { get; set; }

        [DataMember(Order = 2)]
        public string Name { get; set; }

        [DataMember(Order = 3)]
        public string DataType { get; set; }

        [DataMember(Order = 4)]
        public int ItemCount { get; set; }

        [DataMember(Order = 5)]
        public int DistinctValueCount { get; set; }

        [DataMember(Order = 6)]
        public List<string> SampleValues { get; set; } = new();
    }

    [DataContract]
    internal class ScrapeSession
    {
        private List<RawEntry> _rawEntries = new();

        [DataMember(Order = 1)]
        public Guid Id { get; set; } = Guid.NewGuid();

        [DataMember(Order = 2)]
        public DateTime Timestamp { get; set; } = DateTime.Now;

        [DataMember(Order = 3)]
        public string ProfileName { get; set; }

        [DataMember(Order = 4)]
        public string ScopeType { get; set; }

        [DataMember(Order = 5)]
        public string ScopeDescription { get; set; }

        [DataMember(Order = 6)]
        public int ItemsScanned { get; set; }

        [DataMember(Order = 7)]
        public string DocumentFile { get; set; }

        [DataMember(Order = 8)]
        public string DocumentFileKey { get; set; }

        [DataMember(Order = 9)]
        public bool RawEntriesTruncated { get; set; }

        [DataMember(Order = 10)]
        public List<ScrapedProperty> Properties { get; set; } = new();

        [DataMember(Order = 11)]
        public int RawEntryCount { get; set; }

        [DataMember(Order = 12)]
        public string RawEntriesFileName { get; set; }

        [IgnoreDataMember]
        internal bool RawEntriesLoaded { get; set; } = true;

        [IgnoreDataMember]
        public List<RawEntry> RawEntries
        {
            get
            {
                if (RawEntriesLoaded)
                {
                    _rawEntries ??= new List<RawEntry>();
                    return _rawEntries;
                }

                return DataScraperCache.GetOrLoadRawEntries(this);
            }
            set
            {
                _rawEntries = value ?? new List<RawEntry>();
                RawEntryCount = _rawEntries.Count;
                RawEntriesLoaded = true;
            }
        }

        [IgnoreDataMember]
        internal List<RawEntry> RawEntriesUnsafe
        {
            get => _rawEntries ??= new List<RawEntry>();
            set => _rawEntries = value ?? new List<RawEntry>();
        }
    }

    [DataContract]
    internal class RawEntry
    {
        [DataMember(Order = 1)]
        public string Profile { get; set; }

        [DataMember(Order = 2)]
        public string Scope { get; set; }

        [DataMember(Order = 3)]
        public string ItemKey { get; set; }

        [DataMember(Order = 4)]
        public string ItemPath { get; set; }

        [DataMember(Order = 5)]
        public string Category { get; set; }

        [DataMember(Order = 6)]
        public string Name { get; set; }

        [DataMember(Order = 7)]
        public string DataType { get; set; }

        [DataMember(Order = 8)]
        public string Value { get; set; }
    }

    [DataContract]
    internal sealed class ScrapeSessionRecord
    {
        [DataMember(Order = 1)] public Guid Id { get; set; } = Guid.NewGuid();
        [DataMember(Order = 2)] public DateTime Timestamp { get; set; } = DateTime.Now;
        [DataMember(Order = 3)] public string ProfileName { get; set; }
        [DataMember(Order = 4)] public string ScopeType { get; set; }
        [DataMember(Order = 5)] public string ScopeDescription { get; set; }
        [DataMember(Order = 6)] public int ItemsScanned { get; set; }
        [DataMember(Order = 7)] public string DocumentFile { get; set; }
        [DataMember(Order = 8)] public string DocumentFileKey { get; set; }
        [DataMember(Order = 9)] public bool RawEntriesTruncated { get; set; }
        [DataMember(Order = 10)] public List<ScrapedProperty> Properties { get; set; } = new List<ScrapedProperty>();
        [DataMember(Order = 11)] public int RawEntryCount { get; set; }
        [DataMember(Order = 12)] public string RawEntriesFileName { get; set; }

        // Legacy compatibility: old store format had all raw rows inline in ScrapeSessions.json.
        [DataMember(Order = 13, EmitDefaultValue = false)]
        public List<RawEntry> RawEntries { get; set; }
    }

    [DataContract]
    internal sealed class DataScraperCacheStore
    {
        [DataMember(Order = 1)] public int Version { get; set; } = 2;
        [DataMember(Order = 2)] public Guid? LastSessionId { get; set; }
        [DataMember(Order = 3)] public List<ScrapeSessionRecord> Sessions { get; set; } = new List<ScrapeSessionRecord>();
    }

    internal static class DataScraperCache
    {
        private const string StoreFileName = "ScrapeSessions.json";
        private const string RawEntriesDirectoryName = "DataScraperRaw";

        private static readonly object Gate = new object();
        private static readonly List<ScrapeSession> Sessions = new();
        private static ScrapeSession _lastSession;
        private static bool _isLoaded;
        private static string _loadedPath;

        static DataScraperCache()
        {
            MicroEngStorageSettings.Changed += OnStorageSettingsChanged;
        }

        public static ScrapeSession LastSession
        {
            get
            {
                lock (Gate)
                {
                    EnsureLoadedNoLock();
                    return _lastSession;
                }
            }
            set
            {
                lock (Gate)
                {
                    EnsureLoadedNoLock();
                    _lastSession = value;
                }
            }
        }

        public static List<ScrapeSession> AllSessions
        {
            get
            {
                lock (Gate)
                {
                    EnsureLoadedNoLock();
                    return Sessions;
                }
            }
        }

        public static event Action<ScrapeSession> SessionAdded;
        public static event Action CacheChanged;

        internal static List<RawEntry> GetOrLoadRawEntries(ScrapeSession session)
        {
            if (session == null)
            {
                return new List<RawEntry>();
            }

            lock (Gate)
            {
                EnsureLoadedNoLock();
                return EnsureRawEntriesLoadedNoLock(session);
            }
        }

        public static void ReleaseRawEntries(ScrapeSession session)
        {
            if (session == null)
            {
                return;
            }

            lock (Gate)
            {
                EnsureLoadedNoLock();
                if (!session.RawEntriesLoaded)
                {
                    return;
                }

                if (session.RawEntriesUnsafe.Count == 0)
                {
                    session.RawEntriesLoaded = session.RawEntryCount == 0;
                    return;
                }

                session.RawEntriesUnsafe = new List<RawEntry>();
                session.RawEntriesLoaded = session.RawEntryCount == 0;
            }
        }

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
            if (session == null)
            {
                return;
            }

            lock (Gate)
            {
                EnsureLoadedNoLock();

                session.Properties ??= new List<ScrapedProperty>();
                session.RawEntriesUnsafe ??= new List<RawEntry>();
                session.RawEntriesLoaded = true;
                session.RawEntryCount = session.RawEntriesUnsafe.Count;

                SaveRawEntriesNoLock(session);

                Sessions.RemoveAll(s => s != null && s.Id == session.Id);
                Sessions.Add(session);
                _lastSession = session;
                SaveNoLock();
            }

            RaiseSessionAdded(session);
        }

        public static bool RemoveSession(Guid sessionId)
        {
            if (sessionId == Guid.Empty)
            {
                return false;
            }

            lock (Gate)
            {
                EnsureLoadedNoLock();

                var removedSessions = Sessions
                    .Where(s => s != null && s.Id == sessionId)
                    .ToList();
                if (removedSessions.Count == 0)
                {
                    return false;
                }

                foreach (var session in removedSessions)
                {
                    TryDeleteRawEntriesFileNoLock(session);
                }

                Sessions.RemoveAll(s => s != null && s.Id == sessionId);

                if (_lastSession != null && _lastSession.Id == sessionId)
                {
                    _lastSession = Sessions
                        .OrderByDescending(s => s.Timestamp)
                        .FirstOrDefault();
                }

                SaveNoLock();
            }

            RaiseCacheChanged();
            return true;
        }

        public static void ReloadFromStorage()
        {
            lock (Gate)
            {
                _isLoaded = false;
                _loadedPath = null;
                EnsureLoadedNoLock();
            }

            RaiseCacheChanged();
        }

        public static void ClearSessions(bool deletePersistedFile)
        {
            lock (Gate)
            {
                EnsureLoadedNoLock();

                foreach (var session in Sessions.Where(s => s != null))
                {
                    TryDeleteRawEntriesFileNoLock(session);
                }

                Sessions.Clear();
                _lastSession = null;

                if (deletePersistedFile)
                {
                    TryDeleteStoreFileNoLock();
                }
                else
                {
                    SaveNoLock();
                }
            }

            RaiseCacheChanged();
        }

        public static string GetStoreFilePath()
        {
            lock (Gate)
            {
                EnsureLoadedNoLock();
                return _loadedPath ?? string.Empty;
            }
        }

        private static void OnStorageSettingsChanged()
        {
            ReloadFromStorage();
        }

        private static void EnsureLoadedNoLock()
        {
            var path = ResolveStorePathNoLock();
            if (_isLoaded && string.Equals(_loadedPath, path, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            Sessions.Clear();
            _lastSession = null;
            _loadedPath = path;
            var readPath = ResolveReadPath(path);
            var shouldResave = false;

            try
            {
                if (!File.Exists(readPath))
                {
                    _isLoaded = true;
                    return;
                }

                DataScraperCacheStore store;
                using (var fs = File.OpenRead(readPath))
                {
                    var ser = new DataContractJsonSerializer(typeof(DataScraperCacheStore));
                    store = ser.ReadObject(fs) as DataScraperCacheStore ?? new DataScraperCacheStore();
                }

                if (store.Sessions != null && store.Sessions.Count > 0)
                {
                    foreach (var record in store.Sessions.Where(s => s != null))
                    {
                        var session = FromRecordNoLock(record, out var hadInlineRawEntries, out var needsMetadataSave);
                        Sessions.Add(session);

                        if (hadInlineRawEntries)
                        {
                            SaveRawEntriesNoLock(session);
                            session.RawEntriesUnsafe = new List<RawEntry>();
                            session.RawEntriesLoaded = session.RawEntryCount == 0;
                            shouldResave = true;
                        }

                        if (needsMetadataSave)
                        {
                            shouldResave = true;
                        }
                    }
                }

                if (store.LastSessionId.HasValue)
                {
                    _lastSession = Sessions.FirstOrDefault(s => s.Id == store.LastSessionId.Value);
                }

                if (_lastSession == null && Sessions.Count > 0)
                {
                    _lastSession = Sessions
                        .OrderByDescending(s => s.Timestamp)
                        .FirstOrDefault();
                }

                if (!string.Equals(readPath, path, StringComparison.OrdinalIgnoreCase))
                {
                    shouldResave = true;
                }

                if (shouldResave)
                {
                    SaveNoLock();
                }
            }
            catch
            {
                TryBackupCorruptFileNoLock(readPath);
                Sessions.Clear();
                _lastSession = null;
            }

            _isLoaded = true;
        }

        private static ScrapeSession FromRecordNoLock(
            ScrapeSessionRecord record,
            out bool hadInlineRawEntries,
            out bool needsMetadataSave)
        {
            var session = new ScrapeSession
            {
                Id = record.Id == Guid.Empty ? Guid.NewGuid() : record.Id,
                Timestamp = record.Timestamp,
                ProfileName = record.ProfileName,
                ScopeType = record.ScopeType,
                ScopeDescription = record.ScopeDescription,
                ItemsScanned = record.ItemsScanned,
                DocumentFile = record.DocumentFile,
                DocumentFileKey = record.DocumentFileKey,
                RawEntriesTruncated = record.RawEntriesTruncated,
                Properties = record.Properties?.Where(p => p != null).ToList() ?? new List<ScrapedProperty>(),
                RawEntryCount = Math.Max(0, record.RawEntryCount),
                RawEntriesFileName = record.RawEntriesFileName
            };

            var inlineRaw = record.RawEntries?.Where(r => r != null).ToList() ?? new List<RawEntry>();
            hadInlineRawEntries = inlineRaw.Count > 0;
            if (hadInlineRawEntries)
            {
                session.RawEntriesUnsafe = inlineRaw;
                session.RawEntriesLoaded = true;
                session.RawEntryCount = inlineRaw.Count;
            }
            else
            {
                session.RawEntriesUnsafe = new List<RawEntry>();
                session.RawEntriesLoaded = session.RawEntryCount == 0;
            }

            needsMetadataSave = false;
            if (session.RawEntryCount > 0 && string.IsNullOrWhiteSpace(session.RawEntriesFileName))
            {
                EnsureRawEntriesFileNameNoLock(session);
                needsMetadataSave = true;
            }

            return session;
        }

        private static ScrapeSessionRecord ToRecordNoLock(ScrapeSession session)
        {
            session.Properties ??= new List<ScrapedProperty>();
            if (session.RawEntriesLoaded)
            {
                session.RawEntryCount = session.RawEntriesUnsafe.Count;
            }

            var count = Math.Max(0, session.RawEntryCount);
            if (count > 0)
            {
                EnsureRawEntriesFileNameNoLock(session);
            }
            else
            {
                session.RawEntriesFileName = null;
            }

            return new ScrapeSessionRecord
            {
                Id = session.Id,
                Timestamp = session.Timestamp,
                ProfileName = session.ProfileName,
                ScopeType = session.ScopeType,
                ScopeDescription = session.ScopeDescription,
                ItemsScanned = session.ItemsScanned,
                DocumentFile = session.DocumentFile,
                DocumentFileKey = session.DocumentFileKey,
                RawEntriesTruncated = session.RawEntriesTruncated,
                Properties = session.Properties.Where(p => p != null).ToList(),
                RawEntryCount = count,
                RawEntriesFileName = session.RawEntriesFileName
            };
        }

        private static List<RawEntry> EnsureRawEntriesLoadedNoLock(ScrapeSession session)
        {
            if (session == null)
            {
                return new List<RawEntry>();
            }

            if (session.RawEntriesLoaded)
            {
                var loaded = session.RawEntriesUnsafe;
                if (session.RawEntryCount == 0 && loaded.Count > 0)
                {
                    session.RawEntryCount = loaded.Count;
                }

                return loaded;
            }

            var entries = new List<RawEntry>();
            try
            {
                var path = GetRawEntriesPathNoLock(session, ensureFileName: false);
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    using (var fs = File.OpenRead(path))
                    {
                        var ser = new DataContractJsonSerializer(typeof(List<RawEntry>));
                        entries = ser.ReadObject(fs) as List<RawEntry> ?? new List<RawEntry>();
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MicroEngActions.Log($"DataScraperCache raw load failed for session {session.Id}: {ex.Message}");
                }
                catch
                {
                    // ignore
                }
            }

            session.RawEntriesUnsafe = entries;
            session.RawEntriesLoaded = true;
            session.RawEntryCount = entries.Count;
            return entries;
        }

        private static void SaveRawEntriesNoLock(ScrapeSession session)
        {
            if (session == null || !session.RawEntriesLoaded)
            {
                return;
            }

            var entries = session.RawEntriesUnsafe ?? new List<RawEntry>();
            session.RawEntryCount = entries.Count;
            var path = GetRawEntriesPathNoLock(session, ensureFileName: entries.Count > 0);

            if (entries.Count == 0)
            {
                TryDeleteRawEntriesFileNoLock(session);
                session.RawEntriesFileName = null;
                return;
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            var tempPath = path + ".tmp";
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? GetRawEntriesDirectoryPathNoLock());

            try
            {
                using (var fs = File.Create(tempPath))
                {
                    var ser = new DataContractJsonSerializer(typeof(List<RawEntry>));
                    ser.WriteObject(fs, entries);
                }

                if (!ReplaceFileWithRetryNoLock(tempPath, path, out var replaceError))
                {
                    throw replaceError ?? new IOException("Failed to move raw entries temp file.");
                }
            }
            catch (Exception ex)
            {
                try
                {
                    if (File.Exists(tempPath))
                    {
                        File.Delete(tempPath);
                    }
                }
                catch
                {
                    // ignore temp cleanup failures
                }

                try
                {
                    MicroEngActions.Log($"DataScraperCache raw save failed: {ex.Message}");
                }
                catch
                {
                    // ignore
                }
            }
        }

        private static void TryDeleteRawEntriesFileNoLock(ScrapeSession session)
        {
            if (session == null)
            {
                return;
            }

            try
            {
                var path = GetRawEntriesPathNoLock(session, ensureFileName: false);
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    File.Delete(path);
                }

                var fallbackPath = Path.Combine(GetRawEntriesDirectoryPathNoLock(), $"{session.Id:N}.json");
                if (!string.Equals(path, fallbackPath, StringComparison.OrdinalIgnoreCase) && File.Exists(fallbackPath))
                {
                    File.Delete(fallbackPath);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MicroEngActions.Log($"DataScraperCache raw delete failed: {ex.Message}");
                }
                catch
                {
                    // ignore
                }
            }
        }

        private static string ResolveStorePathNoLock()
        {
            return MicroEngStorageSettings.GetDataFilePath(StoreFileName);
        }

        private static string ResolveReadPath(string primaryPath)
        {
            if (File.Exists(primaryPath))
            {
                return primaryPath;
            }

            try
            {
                var legacyPath = Path.Combine(
                    MicroEngStorageSettings.LegacyDefaultDataScraperCacheDirectory,
                    StoreFileName);
                if (File.Exists(legacyPath))
                {
                    return legacyPath;
                }
            }
            catch
            {
                // ignore legacy lookup failures
            }

            return primaryPath;
        }

        private static string GetRawEntriesDirectoryPathNoLock()
        {
            return Path.Combine(MicroEngStorageSettings.DataStorageDirectory, RawEntriesDirectoryName);
        }

        private static string EnsureRawEntriesFileNameNoLock(ScrapeSession session)
        {
            if (session == null)
            {
                return string.Empty;
            }

            if (string.IsNullOrWhiteSpace(session.RawEntriesFileName))
            {
                var id = session.Id == Guid.Empty ? Guid.NewGuid() : session.Id;
                session.Id = id;
                session.RawEntriesFileName = $"{id:N}.json";
            }

            return session.RawEntriesFileName;
        }

        private static string GetRawEntriesPathNoLock(ScrapeSession session, bool ensureFileName)
        {
            if (session == null)
            {
                return string.Empty;
            }

            var fileName = session.RawEntriesFileName;
            if (string.IsNullOrWhiteSpace(fileName))
            {
                if (!ensureFileName)
                {
                    return string.Empty;
                }

                fileName = EnsureRawEntriesFileNameNoLock(session);
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                return string.Empty;
            }

            return Path.Combine(GetRawEntriesDirectoryPathNoLock(), fileName);
        }

        private static void SaveNoLock()
        {
            var path = _loadedPath ?? ResolveStorePathNoLock();
            var tempPath = path + ".tmp";
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? MicroEngStorageSettings.DataStorageDirectory);

            var store = new DataScraperCacheStore
            {
                LastSessionId = _lastSession?.Id,
                Sessions = Sessions
                    .Where(s => s != null)
                    .Select(ToRecordNoLock)
                    .ToList()
            };

            try
            {
                using (var fs = File.Create(tempPath))
                {
                    var ser = new DataContractJsonSerializer(typeof(DataScraperCacheStore));
                    ser.WriteObject(fs, store);
                }

                if (!ReplaceFileWithRetryNoLock(tempPath, path, out var replaceError))
                {
                    throw replaceError ?? new IOException("Failed to move cache store temp file.");
                }
            }
            catch (Exception ex)
            {
                try
                {
                    if (File.Exists(tempPath))
                    {
                        File.Delete(tempPath);
                    }
                }
                catch
                {
                    // ignore temp cleanup failures
                }

                try
                {
                    MicroEngActions.Log($"DataScraperCache save failed: {ex}");
                }
                catch
                {
                    // ignore
                }
            }
        }

        private static bool ReplaceFileWithRetryNoLock(string tempPath, string targetPath, out Exception lastError)
        {
            lastError = null;
            for (var attempt = 0; attempt < 8; attempt++)
            {
                try
                {
                    if (File.Exists(targetPath))
                    {
                        File.Delete(targetPath);
                    }

                    File.Move(tempPath, targetPath);
                    return true;
                }
                catch (IOException ex)
                {
                    lastError = ex;
                }
                catch (UnauthorizedAccessException ex)
                {
                    lastError = ex;
                }

                Thread.Sleep(40 * (attempt + 1));
            }

            return false;
        }

        private static void TryDeleteStoreFileNoLock()
        {
            try
            {
                var path = _loadedPath ?? ResolveStorePathNoLock();
                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                var legacyPath = Path.Combine(
                    MicroEngStorageSettings.LegacyDefaultDataScraperCacheDirectory,
                    StoreFileName);
                if (!string.Equals(path, legacyPath, StringComparison.OrdinalIgnoreCase) && File.Exists(legacyPath))
                {
                    File.Delete(legacyPath);
                }

                var rawDir = GetRawEntriesDirectoryPathNoLock();
                if (Directory.Exists(rawDir))
                {
                    Directory.Delete(rawDir, recursive: true);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MicroEngActions.Log($"DataScraperCache delete failed: {ex.Message}");
                }
                catch
                {
                    // ignore
                }
            }
        }

        private static void TryBackupCorruptFileNoLock(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    return;
                }

                var backup = path + ".corrupt_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".bak";
                File.Copy(path, backup, overwrite: true);
            }
            catch
            {
                // ignore backup failures
            }
        }

        private static void RaiseSessionAdded(ScrapeSession session)
        {
            try
            {
                SessionAdded?.Invoke(session);
            }
            catch
            {
                // ignore subscriber failures
            }
        }

        private static void RaiseCacheChanged()
        {
            try
            {
                CacheChanged?.Invoke();
            }
            catch
            {
                // ignore subscriber failures
            }
        }
    }
}
