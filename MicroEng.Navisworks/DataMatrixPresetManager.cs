using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace MicroEng.Navisworks
{
    internal interface IDataMatrixPresetManager
    {
        IEnumerable<DataMatrixViewPreset> GetPresets(string profileName);
        void SavePreset(DataMatrixViewPreset preset);
        void DeletePreset(string presetId);
    }

    internal class InMemoryPresetManager : IDataMatrixPresetManager
    {
        private readonly List<DataMatrixViewPreset> _presets = new();

        public IEnumerable<DataMatrixViewPreset> GetPresets(string profileName)
        {
            var profile = string.IsNullOrWhiteSpace(profileName) ? "Default" : profileName.Trim();
            return _presets
                .Where(p => string.Equals(p.ScraperProfileName ?? "Default", profile, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        public void SavePreset(DataMatrixViewPreset preset)
        {
            var existing = _presets.FirstOrDefault(p => string.Equals(p.Id, preset.Id, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                _presets.Remove(existing);
            }
            if (string.IsNullOrWhiteSpace(preset.Id))
            {
                preset.Id = Guid.NewGuid().ToString();
            }
            _presets.Add(preset);
        }

        public void DeletePreset(string presetId)
        {
            var existing = _presets.FirstOrDefault(p => string.Equals(p.Id, presetId, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                _presets.Remove(existing);
            }
        }
    }

    [DataContract]
    internal sealed class DataMatrixPresetStore
    {
        [DataMember(Order = 1)] public int Version { get; set; } = 1;
        [DataMember(Order = 2)] public List<DataMatrixViewPreset> Presets { get; set; } = new List<DataMatrixViewPreset>();
    }

    internal sealed class FilePresetManager : IDataMatrixPresetManager
    {
        private readonly object _gate = new object();
        private readonly string _path;

        public FilePresetManager()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MicroEng", "Navisworks", "DataMatrix");

            Directory.CreateDirectory(dir);
            _path = Path.Combine(dir, "ViewPresets.json");
        }

        public IEnumerable<DataMatrixViewPreset> GetPresets(string profileName)
        {
            var profile = string.IsNullOrWhiteSpace(profileName) ? "Default" : profileName.Trim();

            lock (_gate)
            {
                var store = LoadStore();
                return store.Presets
                    .Where(p => string.Equals(p.ScraperProfileName ?? "Default", profile, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(p => p.Name)
                    .ToList();
            }
        }

        public void SavePreset(DataMatrixViewPreset preset)
        {
            if (preset == null) return;

            lock (_gate)
            {
                var store = LoadStore();
                if (string.IsNullOrWhiteSpace(preset.Id))
                {
                    preset.Id = Guid.NewGuid().ToString();
                }

                var existing = store.Presets.FirstOrDefault(p => string.Equals(p.Id, preset.Id, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    store.Presets.Remove(existing);
                }

                store.Presets.Add(preset);
                SaveStore(store);
            }
        }

        public void DeletePreset(string presetId)
        {
            if (string.IsNullOrWhiteSpace(presetId)) return;

            lock (_gate)
            {
                var store = LoadStore();
                store.Presets.RemoveAll(p => string.Equals(p.Id, presetId, StringComparison.OrdinalIgnoreCase));
                SaveStore(store);
            }
        }

        private DataMatrixPresetStore LoadStore()
        {
            try
            {
                if (!File.Exists(_path))
                {
                    return new DataMatrixPresetStore();
                }

                using (var fs = File.OpenRead(_path))
                {
                    var ser = new DataContractJsonSerializer(typeof(DataMatrixPresetStore));
                    var obj = ser.ReadObject(fs) as DataMatrixPresetStore;
                    return obj ?? new DataMatrixPresetStore();
                }
            }
            catch
            {
                try
                {
                    if (File.Exists(_path))
                    {
                        var bak = _path + ".corrupt_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".bak";
                        File.Copy(_path, bak, overwrite: true);
                    }
                }
                catch
                {
                    // ignore
                }

                return new DataMatrixPresetStore();
            }
        }

        private void SaveStore(DataMatrixPresetStore store)
        {
            using (var fs = File.Create(_path))
            {
                var ser = new DataContractJsonSerializer(typeof(DataMatrixPresetStore));
                ser.WriteObject(fs, store ?? new DataMatrixPresetStore());
            }
        }
    }
}
