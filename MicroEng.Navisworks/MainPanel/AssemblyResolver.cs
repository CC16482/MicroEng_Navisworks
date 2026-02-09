using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace MicroEng.Navisworks
{
    internal static class AssemblyResolver
    {
        private static readonly object Sync = new object();
        private static bool _registered;
        private static bool _firstChanceHooked;
        private static int _firstChanceLogged;
        private const int FirstChanceLogLimit = 10;

        private static readonly string AssemblyDir =
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? AppDomain.CurrentDomain.BaseDirectory;

        private static readonly string LogFilePathPrimary = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MicroEng.Navisworks",
            "NavisErrors",
            "MicroEng.log");

        private static readonly string LogFilePathFallback = Path.Combine(Path.GetTempPath(), "MicroEng.log");

        private static readonly HashSet<string> ResolveLogOnce = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        [ThreadStatic]
        private static HashSet<string> _resolving;

        public static void EnsureRegistered()
        {
            if (_registered) return;

            lock (Sync)
            {
                if (_registered) return;
                AppDomain.CurrentDomain.AssemblyResolve += OnResolve;
                if (!_firstChanceHooked)
                {
                    AppDomain.CurrentDomain.FirstChanceException += OnFirstChance;
                    _firstChanceHooked = true;
                }
                _registered = true;
            }
        }

        private static void OnFirstChance(object sender, System.Runtime.ExceptionServices.FirstChanceExceptionEventArgs e)
        {
            try
            {
                // Only log the first few FileLoad/FileNotFound exceptions for our assemblies.
                if (_firstChanceLogged >= FirstChanceLogLimit)
                    return;

                var ex = e.Exception;
                switch (ex)
                {
                    case System.IO.FileLoadException fle:
                        if (IsOurs(fle.FileName))
                        {
                            _firstChanceLogged++;
                            SafeLog($"[FirstChance] FileLoad: {fle.FileName} :: {fle.Message}");
                        }
                        break;
                    case System.IO.FileNotFoundException fnf:
                        if (IsOurs(fnf.FileName))
                        {
                            _firstChanceLogged++;
                            SafeLog($"[FirstChance] FileNotFound: {fnf.FileName} :: {fnf.Message}");
                        }
                        break;
                }
            }
            catch
            {
                // swallow
            }
        }

        private static bool IsOurs(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return false;
            return fileName.IndexOf("MicroEng", StringComparison.OrdinalIgnoreCase) >= 0
                   || fileName.IndexOf("Wpf.Ui", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static Assembly OnResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                var requested = new AssemblyName(args.Name);
                var name = requested.Name ?? string.Empty;
                if (string.IsNullOrEmpty(name))
                    return null;

                // Ignore resource satellites.
                if (name.EndsWith(".resources", StringComparison.OrdinalIgnoreCase))
                    return null;

                // Only handle our own and Wpf.Ui assemblies.
                if (!ShouldHandle(name))
                {
                    return null;
                }

                // If it's already loaded, never load it again (avoids duplicate-load issues that can break WPF pack URIs).
                var alreadyLoaded = TryGetLoaded(name);
                if (alreadyLoaded != null)
                {
                    return alreadyLoaded;
                }

                _resolving ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (!_resolving.Add(name))
                {
                    return null;
                }

                try
                {
                    var candidate = Path.Combine(AssemblyDir, name + ".dll");
                    if (!File.Exists(candidate))
                    {
                        return null;
                    }

                    // Re-check in case another thread loaded it while we were waiting.
                    alreadyLoaded = TryGetLoaded(name);
                    if (alreadyLoaded != null)
                    {
                        return alreadyLoaded;
                    }

                    SafeResolveLogOnce($"[AssemblyResolve] Loading {name} from {candidate}");
                    return Assembly.LoadFrom(candidate);
                }
                finally
                {
                    _resolving.Remove(name);
                }
            }
            catch
            {
                // swallow
            }
            return null;
        }

        private static bool ShouldHandle(string assemblySimpleName)
        {
            return assemblySimpleName.StartsWith("MicroEng", StringComparison.OrdinalIgnoreCase)
                   || assemblySimpleName.StartsWith("Wpf.Ui", StringComparison.OrdinalIgnoreCase);
        }

        private static Assembly TryGetLoaded(string simpleName)
        {
            try
            {
                foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    var asmName = asm.GetName().Name;
                    if (string.Equals(asmName, simpleName, StringComparison.OrdinalIgnoreCase))
                    {
                        return asm;
                    }
                }
            }
            catch
            {
                // swallow
            }

            return null;
        }

        private static void SafeResolveLogOnce(string message)
        {
            try
            {
                lock (Sync)
                {
                    if (!ResolveLogOnce.Add(message))
                    {
                        return;
                    }
                }

                SafeLog(message);
            }
            catch
            {
                // swallow
            }
        }

        private static void SafeLog(string message)
        {
            var line = $"{DateTime.Now:HH:mm:ss} {message}";
            WriteLogLine(LogFilePathPrimary, line);
            WriteLogLine(LogFilePathFallback, line);
        }

        private static void WriteLogLine(string path, string line)
        {
            try
            {
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                File.AppendAllLines(path, new[] { line });
            }
            catch
            {
                // swallow
            }
        }
    }
}
