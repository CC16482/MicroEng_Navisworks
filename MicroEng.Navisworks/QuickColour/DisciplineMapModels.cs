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
            if (!File.Exists(path))
            {
                return new DisciplineMapFile();
            }

            var json = ReadAllTextNormalized(path);
            if (string.IsNullOrWhiteSpace(json))
            {
                return new DisciplineMapFile();
            }

            var startObj = json.IndexOf('{');
            var startArr = json.IndexOf('[');
            var start = startObj >= 0 && startArr >= 0 ? Math.Min(startObj, startArr) : Math.Max(startObj, startArr);
            if (start > 0)
            {
                json = json.Substring(start);
            }

            var bytes = Encoding.UTF8.GetBytes(json);
            using (var ms = new MemoryStream(bytes))
            {
                var ser = new DataContractJsonSerializer(typeof(DisciplineMapFile));
                return (DisciplineMapFile)ser.ReadObject(ms);
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
                File.WriteAllText(path, defaultJson ?? "{}", new UTF8Encoding(false));
            }
        }

        private static string ReadAllTextNormalized(string path)
        {
            using (var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
            {
                var text = reader.ReadToEnd();
                if (string.IsNullOrEmpty(text))
                {
                    return text;
                }

                if (text[0] == '\uFEFF')
                {
                    text = text.Substring(1);
                }

                return text.TrimStart();
            }
        }
    }
}
