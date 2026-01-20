using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace MicroEng.Navisworks.QuickColour
{
    [DataContract]
    public sealed class DisciplineMapFile
    {
        [DataMember(Name = "version")] public int Version { get; set; } = 1;
        [DataMember(Name = "fallbackGroup")] public string FallbackGroup { get; set; } = "Other";
        [DataMember(Name = "rules")] public List<DisciplineMapRule> Rules { get; set; } = new List<DisciplineMapRule>();
    }

    [DataContract]
    public sealed class DisciplineMapRule
    {
        [DataMember(Name = "group")] public string Group { get; set; } = "";
        [DataMember(Name = "match")] public List<DisciplineMapMatcher> Match { get; set; } = new List<DisciplineMapMatcher>();
    }

    [DataContract]
    public sealed class DisciplineMapMatcher
    {
        [DataMember(Name = "type")] public string Type { get; set; } = "exact";
        [DataMember(Name = "value")] public string Value { get; set; } = "";
    }

    internal static class DisciplineMapFileIO
    {
        public static DisciplineMapFile Load(string path)
        {
            using (var fs = File.OpenRead(path))
            {
                var ser = new DataContractJsonSerializer(typeof(DisciplineMapFile));
                return (DisciplineMapFile)ser.ReadObject(fs);
            }
        }

        public static void EnsureDefaultExists(string path, string defaultJson)
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (!File.Exists(path))
            {
                File.WriteAllText(path, defaultJson ?? "{}", Encoding.UTF8);
            }
        }
    }
}
