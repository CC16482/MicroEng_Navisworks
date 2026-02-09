using System;
using System.IO;
using System.Runtime.Serialization.Json;

namespace MicroEng.Navisworks.TreeMapper
{
    internal sealed class TreeMapperPublishStore
    {
        private readonly object _gate = new object();
        private readonly string _activeProfilePath;
        private readonly string _publishedTreePath;
        private readonly string _legacyDir;
        private readonly string _legacyActiveProfilePath;
        private readonly string _legacyPublishedTreePath;

        public TreeMapperPublishStore()
        {
            _activeProfilePath = MicroEngStorageSettings.GetDataFilePath("TreeMapperActiveProfile.json");
            _publishedTreePath = MicroEngStorageSettings.GetDataFilePath("TreeMapperPublishedTree.json");

            _legacyDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MicroEng",
                "Navisworks",
                "TreeMapper");
            _legacyActiveProfilePath = Path.Combine(_legacyDir, "ActiveProfile.json");
            _legacyPublishedTreePath = Path.Combine(_legacyDir, "PublishedTree.json");

            Directory.CreateDirectory(Path.GetDirectoryName(_activeProfilePath) ?? MicroEngStorageSettings.DataStorageDirectory);
            Directory.CreateDirectory(_legacyDir);
        }

        public string PublishDirectory => Path.GetDirectoryName(_publishedTreePath) ?? MicroEngStorageSettings.DataStorageDirectory;
        public string ActiveProfilePath => _activeProfilePath;
        public string PublishedTreePath => _publishedTreePath;

        public void SaveActiveProfile(TreeMapperProfile profile)
        {
            lock (_gate)
            {
                using (var fs = File.Create(_activeProfilePath))
                {
                    var ser = new DataContractJsonSerializer(typeof(TreeMapperProfile));
                    ser.WriteObject(fs, profile ?? new TreeMapperProfile());
                }

                // Backward compatibility for existing external readers.
                using (var fs = File.Create(_legacyActiveProfilePath))
                {
                    var ser = new DataContractJsonSerializer(typeof(TreeMapperProfile));
                    ser.WriteObject(fs, profile ?? new TreeMapperProfile());
                }
            }
        }

        public void SavePublishedTree(TreeMapperPublishedTree tree)
        {
            lock (_gate)
            {
                using (var fs = File.Create(_publishedTreePath))
                {
                    var ser = new DataContractJsonSerializer(typeof(TreeMapperPublishedTree));
                    ser.WriteObject(fs, tree ?? new TreeMapperPublishedTree());
                }

                // Backward compatibility for existing external readers.
                using (var fs = File.Create(_legacyPublishedTreePath))
                {
                    var ser = new DataContractJsonSerializer(typeof(TreeMapperPublishedTree));
                    ser.WriteObject(fs, tree ?? new TreeMapperPublishedTree());
                }
            }
        }
    }
}
