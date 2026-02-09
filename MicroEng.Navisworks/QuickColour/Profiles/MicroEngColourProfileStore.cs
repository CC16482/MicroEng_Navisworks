using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace MicroEng.Navisworks.QuickColour.Profiles
{
    internal sealed class MicroEngColourProfileStoreFile
    {
        [JsonProperty("version")]
        public int Version { get; set; } = 1;

        [JsonProperty("profiles")]
        public List<MicroEngColourProfile> Profiles { get; set; } = new List<MicroEngColourProfile>();
    }

    internal sealed class MicroEngColourProfileStore
    {
        private readonly string _storeFilePath;
        private readonly string _legacyProfilesFolder;

        private static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            MissingMemberHandling = MissingMemberHandling.Ignore,
            NullValueHandling = NullValueHandling.Include,
            DateParseHandling = DateParseHandling.DateTime,
        };

        public MicroEngColourProfileStore(string storeFilePath, string legacyProfilesFolder = null)
        {
            if (string.IsNullOrWhiteSpace(storeFilePath))
            {
                throw new ArgumentNullException(nameof(storeFilePath));
            }

            _storeFilePath = Path.GetFullPath(storeFilePath);
            _legacyProfilesFolder = string.IsNullOrWhiteSpace(legacyProfilesFolder)
                ? string.Empty
                : Path.GetFullPath(legacyProfilesFolder);

            Directory.CreateDirectory(Path.GetDirectoryName(_storeFilePath) ?? MicroEngStorageSettings.DataStorageDirectory);
        }

        public static MicroEngColourProfileStore CreateDefault()
        {
            var newPath = MicroEngStorageSettings.GetDataFilePath("QuickColourProfiles.json");
            var legacyRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MicroEng",
                "ColourProfiles");

            return new MicroEngColourProfileStore(newPath, legacyRoot);
        }

        public IReadOnlyList<MicroEngColourProfile> LoadAll()
        {
            var store = LoadStore();
            return (store?.Profiles ?? new List<MicroEngColourProfile>())
                .OrderByDescending(p => p.ModifiedUtc)
                .ThenBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public void Save(MicroEngColourProfile profile)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));

            NormalizeProfile(profile, updateModifiedUtc: true);

            var store = LoadStore();
            store.Profiles ??= new List<MicroEngColourProfile>();
            store.Profiles.RemoveAll(p => string.Equals(p.Id, profile.Id, StringComparison.OrdinalIgnoreCase));
            store.Profiles.Add(profile);
            SaveStore(store);
        }

        public void Delete(string profileId)
        {
            if (string.IsNullOrWhiteSpace(profileId))
            {
                return;
            }

            var store = LoadStore();
            store.Profiles ??= new List<MicroEngColourProfile>();
            store.Profiles.RemoveAll(p => string.Equals(p.Id, profileId, StringComparison.OrdinalIgnoreCase));
            SaveStore(store);
        }

        public void ExportTo(MicroEngColourProfile profile, string destinationPath)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (string.IsNullOrWhiteSpace(destinationPath)) throw new ArgumentNullException(nameof(destinationPath));

            profile.SchemaVersion = MicroEngColourProfileSchema.CurrentSchemaVersion;
            var json = JsonConvert.SerializeObject(profile, JsonSettings);
            File.WriteAllText(destinationPath, json);
        }

        public MicroEngColourProfile ImportFrom(string sourcePath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));

            var json = File.ReadAllText(sourcePath);
            var profile = JsonConvert.DeserializeObject<MicroEngColourProfile>(json, JsonSettings);
            if (profile == null) throw new InvalidDataException("Invalid profile JSON.");

            profile = MigrateToCurrent(profile);

            if (string.IsNullOrWhiteSpace(profile.Id))
            {
                profile.Id = Guid.NewGuid().ToString("N");
            }

            profile.SchemaVersion = MicroEngColourProfileSchema.CurrentSchemaVersion;
            profile.ModifiedUtc = DateTime.UtcNow;
            if (profile.CreatedUtc == default(DateTime)) profile.CreatedUtc = DateTime.UtcNow;

            Save(profile);
            return profile;
        }

        private MicroEngColourProfileStoreFile LoadStore()
        {
            try
            {
                if (File.Exists(_storeFilePath))
                {
                    var json = File.ReadAllText(_storeFilePath);
                    var loaded = JsonConvert.DeserializeObject<MicroEngColourProfileStoreFile>(json, JsonSettings)
                                 ?? new MicroEngColourProfileStoreFile();
                    loaded.Profiles = NormalizeProfiles(loaded.Profiles);
                    return loaded;
                }

                var legacyProfiles = LoadLegacyProfiles();
                if (legacyProfiles.Count > 0)
                {
                    var migrated = new MicroEngColourProfileStoreFile { Profiles = legacyProfiles };
                    SaveStore(migrated);
                    return migrated;
                }
            }
            catch
            {
                TryBackupCorruptStore();
            }

            return new MicroEngColourProfileStoreFile();
        }

        private void SaveStore(MicroEngColourProfileStoreFile store)
        {
            var normalized = new MicroEngColourProfileStoreFile
            {
                Version = Math.Max(1, store?.Version ?? 1),
                Profiles = NormalizeProfiles(store?.Profiles)
            };

            Directory.CreateDirectory(Path.GetDirectoryName(_storeFilePath) ?? MicroEngStorageSettings.DataStorageDirectory);

            var tempPath = _storeFilePath + ".tmp";
            var json = JsonConvert.SerializeObject(normalized, JsonSettings);
            File.WriteAllText(tempPath, json);

            if (File.Exists(_storeFilePath))
            {
                File.Delete(_storeFilePath);
            }

            File.Move(tempPath, _storeFilePath);
        }

        private List<MicroEngColourProfile> LoadLegacyProfiles()
        {
            var list = new List<MicroEngColourProfile>();
            if (string.IsNullOrWhiteSpace(_legacyProfilesFolder) || !Directory.Exists(_legacyProfilesFolder))
            {
                return list;
            }

            var files = Directory.EnumerateFiles(_legacyProfilesFolder, "*.json", SearchOption.TopDirectoryOnly).ToList();
            foreach (var file in files)
            {
                try
                {
                    var json = File.ReadAllText(file);
                    var profile = JsonConvert.DeserializeObject<MicroEngColourProfile>(json, JsonSettings);
                    if (profile == null)
                    {
                        continue;
                    }

                    NormalizeProfile(profile, updateModifiedUtc: false);
                    list.Add(profile);
                }
                catch
                {
                    // ignore bad legacy files
                }
            }

            return list;
        }

        private List<MicroEngColourProfile> NormalizeProfiles(IEnumerable<MicroEngColourProfile> profiles)
        {
            var list = new List<MicroEngColourProfile>();
            foreach (var profile in profiles ?? Enumerable.Empty<MicroEngColourProfile>())
            {
                if (profile == null)
                {
                    continue;
                }

                NormalizeProfile(profile, updateModifiedUtc: false);
                list.Add(profile);
            }

            return list;
        }

        private static void NormalizeProfile(MicroEngColourProfile profile, bool updateModifiedUtc)
        {
            if (profile == null)
            {
                return;
            }

            profile.SchemaVersion = MicroEngColourProfileSchema.CurrentSchemaVersion;
            if (updateModifiedUtc)
            {
                profile.ModifiedUtc = DateTime.UtcNow;
            }
            else if (profile.ModifiedUtc == default(DateTime))
            {
                profile.ModifiedUtc = DateTime.UtcNow;
            }
            if (profile.CreatedUtc == default(DateTime))
            {
                profile.CreatedUtc = DateTime.UtcNow;
            }

            if (string.IsNullOrWhiteSpace(profile.Id))
            {
                profile.Id = Guid.NewGuid().ToString("N");
            }

            profile.Generator ??= new MicroEngColourGenerator();
            profile.Scope ??= new MicroEngColourScope();
            profile.Outputs ??= new MicroEngColourOutputs();
            profile.Rules ??= new List<MicroEngColourRule>();
        }

        private static MicroEngColourProfile MigrateToCurrent(MicroEngColourProfile profile)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));

            if (profile.SchemaVersion <= 0)
            {
                profile.SchemaVersion = MicroEngColourProfileSchema.CurrentSchemaVersion;
            }

            return profile;
        }

        private void TryBackupCorruptStore()
        {
            try
            {
                if (!File.Exists(_storeFilePath))
                {
                    return;
                }

                var backup = _storeFilePath + ".corrupt_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".bak";
                File.Copy(_storeFilePath, backup, overwrite: true);
            }
            catch
            {
                // ignore
            }
        }
    }
}
