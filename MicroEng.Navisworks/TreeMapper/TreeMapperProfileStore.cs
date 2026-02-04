using System;
using System.IO;
using System.Runtime.Serialization.Json;

namespace MicroEng.Navisworks.TreeMapper
{
    internal sealed class TreeMapperProfileStore
    {
        private readonly object _gate = new object();
        private readonly string _path;

        public TreeMapperProfileStore()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MicroEng", "Navisworks", "TreeMapper");
            Directory.CreateDirectory(dir);
            _path = Path.Combine(dir, "Profiles.json");
        }

        public TreeMapperProfileStoreData Load()
        {
            lock (_gate)
            {
                try
                {
                    if (!File.Exists(_path))
                    {
                        return new TreeMapperProfileStoreData();
                    }

                    using (var fs = File.OpenRead(_path))
                    {
                        var ser = new DataContractJsonSerializer(typeof(TreeMapperProfileStoreData));
                        var obj = ser.ReadObject(fs) as TreeMapperProfileStoreData;
                        return obj ?? new TreeMapperProfileStoreData();
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
                using (var fs = File.Create(_path))
                {
                    var ser = new DataContractJsonSerializer(typeof(TreeMapperProfileStoreData));
                    ser.WriteObject(fs, store ?? new TreeMapperProfileStoreData());
                }
            }
        }
    }
}
