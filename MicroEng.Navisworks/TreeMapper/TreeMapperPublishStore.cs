using System;
using System.IO;
using System.Runtime.Serialization.Json;

namespace MicroEng.Navisworks.TreeMapper
{
    internal sealed class TreeMapperPublishStore
    {
        private readonly object _gate = new object();
        private readonly string _dir;
        private readonly string _activeProfilePath;
        private readonly string _publishedTreePath;

        public TreeMapperPublishStore()
        {
            _dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MicroEng", "Navisworks", "TreeMapper");
            Directory.CreateDirectory(_dir);
            _activeProfilePath = Path.Combine(_dir, "ActiveProfile.json");
            _publishedTreePath = Path.Combine(_dir, "PublishedTree.json");
        }

        public string PublishDirectory => _dir;
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
            }
        }
    }
}
