using System;
using System.Collections.Generic;
using System.Linq;

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
            return _presets.Where(p => string.Equals(p.ScraperProfileName, profileName, StringComparison.OrdinalIgnoreCase)).ToList();
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
}
