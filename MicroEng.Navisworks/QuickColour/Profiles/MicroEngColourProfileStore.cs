using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace MicroEng.Navisworks.QuickColour.Profiles
{
    internal sealed class MicroEngColourProfileStore
    {
        private readonly string _profilesFolder;

        private static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            MissingMemberHandling = MissingMemberHandling.Ignore,
            NullValueHandling = NullValueHandling.Include,
            DateParseHandling = DateParseHandling.DateTime,
        };

        public MicroEngColourProfileStore(string profilesFolder)
        {
            _profilesFolder = profilesFolder ?? throw new ArgumentNullException(nameof(profilesFolder));
            Directory.CreateDirectory(_profilesFolder);
        }

        public static MicroEngColourProfileStore CreateDefault()
        {
            var root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MicroEng",
                "ColourProfiles");

            return new MicroEngColourProfileStore(root);
        }

        public IReadOnlyList<MicroEngColourProfile> LoadAll()
        {
            Directory.CreateDirectory(_profilesFolder);

            var files = Directory.EnumerateFiles(_profilesFolder, "*.json", SearchOption.TopDirectoryOnly).ToList();
            var list = new List<MicroEngColourProfile>();

            foreach (var f in files)
            {
                try
                {
                    var json = File.ReadAllText(f);
                    var profile = JsonConvert.DeserializeObject<MicroEngColourProfile>(json, JsonSettings);
                    if (profile == null)
                    {
                        continue;
                    }

                    profile = MigrateToCurrent(profile);

                    if (profile.SchemaVersion != MicroEngColourProfileSchema.CurrentSchemaVersion)
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(profile.Id))
                    {
                        continue;
                    }

                    profile.Generator ??= new MicroEngColourGenerator();
                    profile.Scope ??= new MicroEngColourScope();
                    profile.Outputs ??= new MicroEngColourOutputs();
                    profile.Rules ??= new List<MicroEngColourRule>();

                    list.Add(profile);
                }
                catch
                {
                    // ignore bad files
                }
            }

            return list
                .OrderByDescending(p => p.ModifiedUtc)
                .ThenBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public void Save(MicroEngColourProfile profile)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));

            Directory.CreateDirectory(_profilesFolder);

            profile.SchemaVersion = MicroEngColourProfileSchema.CurrentSchemaVersion;
            profile.ModifiedUtc = DateTime.UtcNow;
            if (profile.CreatedUtc == default(DateTime)) profile.CreatedUtc = DateTime.UtcNow;
            if (string.IsNullOrWhiteSpace(profile.Id)) profile.Id = Guid.NewGuid().ToString("N");
            profile.Generator ??= new MicroEngColourGenerator();
            profile.Scope ??= new MicroEngColourScope();
            profile.Outputs ??= new MicroEngColourOutputs();
            profile.Rules ??= new List<MicroEngColourRule>();

            var safeName = MakeSafeFileName(profile.Name);
            var fileName = $"{safeName}_{profile.Id}.json";
            var path = Path.Combine(_profilesFolder, fileName);

            var tmp = path + ".tmp";
            var json = JsonConvert.SerializeObject(profile, JsonSettings);
            File.WriteAllText(tmp, json);

            if (File.Exists(path)) File.Delete(path);
            File.Move(tmp, path);
        }

        public void Delete(string profileId)
        {
            if (string.IsNullOrWhiteSpace(profileId)) return;

            foreach (var f in Directory.EnumerateFiles(_profilesFolder, "*.json"))
            {
                if (f.IndexOf(profileId, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    try { File.Delete(f); } catch { /* ignore */ }
                }
            }
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

        private static MicroEngColourProfile MigrateToCurrent(MicroEngColourProfile profile)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));

            if (profile.SchemaVersion <= 0)
            {
                profile.SchemaVersion = MicroEngColourProfileSchema.CurrentSchemaVersion;
            }

            return profile;
        }

        private static string MakeSafeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Profile";
            var invalid = Path.GetInvalidFileNameChars();
            var cleaned = new string(name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
            cleaned = cleaned.Trim();
            return string.IsNullOrWhiteSpace(cleaned) ? "Profile" : cleaned;
        }
    }
}
