using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace MicroEng.Navisworks
{
    [DataContract]
    internal sealed class MicroEngStorageSettingsModel
    {
        [DataMember(Order = 1)] public int Version { get; set; } = 1;
        [DataMember(Order = 2)] public string DataStorageDirectory { get; set; } = string.Empty;
        [DataMember(Order = 3)] public string DataScraperCacheDirectory { get; set; } = string.Empty;
    }

    internal static class MicroEngStorageSettings
    {
        private static readonly object Gate = new object();
        private static readonly string RootDirectoryPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MicroEng",
            "Navisworks");
        private static readonly string SettingsFilePathValue = Path.Combine(RootDirectoryPath, "Settings.json");
        private static MicroEngStorageSettingsModel _current;

        public static event Action Changed;

        public static string SettingsFilePath => SettingsFilePathValue;
        public static string RootDirectory => RootDirectoryPath;
        public static string DefaultDataStorageDirectory => Path.Combine(RootDirectoryPath, "DataStore");
        public static string LegacyDefaultDataScraperCacheDirectory => Path.Combine(RootDirectoryPath, "DataScraperCache");

        public static string DataStorageDirectory
        {
            get
            {
                lock (Gate)
                {
                    EnsureLoadedNoLock();
                    return ResolveDataDirectoryNoLock();
                }
            }
        }

        public static string DataScraperCacheDirectory => DataStorageDirectory;

        public static bool SetDataStorageDirectory(string directoryPath, out string resolvedDirectoryPath)
        {
            var normalized = NormalizeDirectoryPath(directoryPath);
            var changed = false;

            lock (Gate)
            {
                EnsureLoadedNoLock();
                var previous = ResolveDataDirectoryNoLock();
                _current.DataStorageDirectory = normalized;
                resolvedDirectoryPath = ResolveDataDirectoryNoLock();
                changed = !string.Equals(previous, resolvedDirectoryPath, StringComparison.OrdinalIgnoreCase);
                SaveNoLock();
            }

            if (changed)
            {
                RaiseChanged();
            }

            return changed;
        }

        public static bool SetDataScraperCacheDirectory(string directoryPath, out string resolvedDirectoryPath)
        {
            return SetDataStorageDirectory(directoryPath, out resolvedDirectoryPath);
        }

        public static bool ResetDataStorageDirectory(out string resolvedDirectoryPath)
        {
            return SetDataStorageDirectory(string.Empty, out resolvedDirectoryPath);
        }

        public static bool ResetDataScraperCacheDirectory(out string resolvedDirectoryPath)
        {
            return ResetDataStorageDirectory(out resolvedDirectoryPath);
        }

        public static string GetDataFilePath(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException("Invalid file name.", nameof(fileName));
            }

            var safeName = fileName.Trim();
            if (safeName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                throw new ArgumentException("Invalid file name.", nameof(fileName));
            }

            return Path.Combine(DataStorageDirectory, safeName);
        }

        public static string GetDataSubdirectory(string subdirectoryName)
        {
            if (string.IsNullOrWhiteSpace(subdirectoryName))
            {
                throw new ArgumentException("Invalid subdirectory name.", nameof(subdirectoryName));
            }

            var dir = Path.Combine(DataStorageDirectory, subdirectoryName.Trim());
            Directory.CreateDirectory(dir);
            return dir;
        }

        private static void EnsureLoadedNoLock()
        {
            if (_current != null)
            {
                return;
            }

            _current = LoadNoLock();
        }

        private static MicroEngStorageSettingsModel LoadNoLock()
        {
            try
            {
                if (!File.Exists(SettingsFilePathValue))
                {
                    return new MicroEngStorageSettingsModel();
                }

                using (var fs = File.OpenRead(SettingsFilePathValue))
                {
                    var ser = new DataContractJsonSerializer(typeof(MicroEngStorageSettingsModel));
                    var loaded = ser.ReadObject(fs) as MicroEngStorageSettingsModel;
                    return loaded ?? new MicroEngStorageSettingsModel();
                }
            }
            catch
            {
                try
                {
                    if (File.Exists(SettingsFilePathValue))
                    {
                        var backup = SettingsFilePathValue + ".corrupt_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".bak";
                        File.Copy(SettingsFilePathValue, backup, overwrite: true);
                    }
                }
                catch
                {
                    // ignore
                }

                return new MicroEngStorageSettingsModel();
            }
        }

        private static void SaveNoLock()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsFilePathValue) ?? RootDirectoryPath);
                using (var fs = File.Create(SettingsFilePathValue))
                {
                    var ser = new DataContractJsonSerializer(typeof(MicroEngStorageSettingsModel));
                    ser.WriteObject(fs, _current ?? new MicroEngStorageSettingsModel());
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MicroEngActions.Log($"Storage settings save failed: {ex.Message}");
                }
                catch
                {
                    // ignore logging failures
                }
            }
        }

        private static string ResolveDataDirectoryNoLock()
        {
            var configured = NormalizeDirectoryPath(_current?.DataStorageDirectory);
            if (string.IsNullOrWhiteSpace(configured))
            {
                configured = NormalizeDirectoryPath(_current?.DataScraperCacheDirectory);
            }

            if (string.IsNullOrWhiteSpace(configured))
            {
                configured = DefaultDataStorageDirectory;
            }

            Directory.CreateDirectory(configured);
            return configured;
        }

        private static string NormalizeDirectoryPath(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            try
            {
                var expanded = Environment.ExpandEnvironmentVariables(input.Trim());
                if (string.IsNullOrWhiteSpace(expanded))
                {
                    return string.Empty;
                }

                if (!Path.IsPathRooted(expanded))
                {
                    expanded = Path.GetFullPath(Path.Combine(RootDirectoryPath, expanded));
                }
                else
                {
                    expanded = Path.GetFullPath(expanded);
                }

                return expanded.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void RaiseChanged()
        {
            try
            {
                Changed?.Invoke();
            }
            catch
            {
                // never let listeners break settings updates
            }
        }
    }
}
