using System;
using System.IO;
using System.Runtime.Serialization.Json;

namespace MicroEng.Navisworks.TreeMapper
{
    internal sealed class TreeMapperProfileStore
    {
        private readonly object _gate = new object();
        private readonly string _path;
        private readonly string _legacyPath;
        private bool _legacyMigrated;

        public TreeMapperProfileStore()
        {
            _path = MicroEngStorageSettings.GetDataFilePath("TreeMapperProfiles.json");
            _legacyPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MicroEng",
                "Navisworks",
                "TreeMapper",
                "Profiles.json");
            Directory.CreateDirectory(Path.GetDirectoryName(_path) ?? MicroEngStorageSettings.DataStorageDirectory);
        }

        public TreeMapperProfileStoreData Load()
        {
            lock (_gate)
            {
                try
                {
                    var readPath = ResolveReadPath();
                    if (!File.Exists(readPath))
                    {
                        return new TreeMapperProfileStoreData();
                    }

                    using (var fs = File.OpenRead(readPath))
                    {
                        var ser = new DataContractJsonSerializer(typeof(TreeMapperProfileStoreData));
                        var obj = ser.ReadObject(fs) as TreeMapperProfileStoreData;
                        var result = obj ?? new TreeMapperProfileStoreData();
                        TryMigrateLegacy(readPath, result);
                        return result;
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

                    return new TreeMapperProfileStoreData();
                }
            }
        }

        public void Save(TreeMapperProfileStoreData store)
        {
            lock (_gate)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_path) ?? MicroEngStorageSettings.DataStorageDirectory);
                using (var fs = File.Create(_path))
                {
                    var ser = new DataContractJsonSerializer(typeof(TreeMapperProfileStoreData));
                    ser.WriteObject(fs, store ?? new TreeMapperProfileStoreData());
                }
            }
        }

        private string ResolveReadPath()
        {
            if (File.Exists(_path))
            {
                return _path;
            }

            if (File.Exists(_legacyPath))
            {
                return _legacyPath;
            }

            return _path;
        }

        private void TryMigrateLegacy(string readPath, TreeMapperProfileStoreData store)
        {
            if (_legacyMigrated)
            {
                return;
            }

            if (string.Equals(readPath, _legacyPath, StringComparison.OrdinalIgnoreCase))
            {
                Save(store ?? new TreeMapperProfileStoreData());
                _legacyMigrated = true;
            }
        }
    }
}
