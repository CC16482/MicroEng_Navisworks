using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

namespace MicroEng.Navisworks.SelectionTreeCom
{
    [Plugin("MicroEng.Dev.ComPluginListProbe", "MENG",
        DisplayName = "MicroEng Dev: COM Plugin List Probe",
        ToolTip = "Logs COM plugin records (filtered) to help identify Selection Tree-related plugins.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public sealed class DevComPluginListProbe : AddInPlugin
    {
        private static readonly string[] DefaultNeedles =
        {
            "tree",
            "selection",
            "select",
            "sel"
        };

        public override int Execute(params string[] parameters)
        {
            var logAll = parameters != null
                && parameters.Any(p => string.Equals(p, "all", StringComparison.OrdinalIgnoreCase));
            return Run(logAll, "ComPluginListProbe");
        }

        private static bool MatchesNeedles(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return DefaultNeedles.Any(n => value.IndexOf(n, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static bool TryGetParameter(ComApi.InwPlugin plugin, object pluginObj, int index, out string name, out object value)
        {
            name = null;
            value = null;

            if (plugin == null)
            {
                return TryGetParameterLate(pluginObj, index, out name, out value);
            }

            try
            {
                object data = null;
                name = plugin.iGetParameter(index, ref data);
                value = data;
                return true;
            }
            catch
            {
                return TryGetParameterLate(pluginObj, index, out name, out value);
            }
        }

        private static bool TryGetParameterLate(object pluginObj, int index, out string name, out object value)
        {
            name = null;
            value = null;

            if (pluginObj == null)
            {
                return false;
            }

            try
            {
                object data = null;
                name = (string)pluginObj.GetType().InvokeMember(
                    "iGetParameter",
                    System.Reflection.BindingFlags.InvokeMethod
                    | System.Reflection.BindingFlags.Public
                    | System.Reflection.BindingFlags.Instance,
                    null,
                    pluginObj,
                    new object[] { index, data });
                value = data;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string TryGetStringMember(object obj, string memberName)
        {
            if (obj == null || string.IsNullOrWhiteSpace(memberName))
            {
                return null;
            }

            try
            {
                var value = obj.GetType().InvokeMember(
                    memberName,
                    System.Reflection.BindingFlags.GetProperty
                    | System.Reflection.BindingFlags.InvokeMethod
                    | System.Reflection.BindingFlags.Public
                    | System.Reflection.BindingFlags.Instance,
                    null,
                    obj,
                    Array.Empty<object>());
                return value as string;
            }
            catch
            {
                return null;
            }
        }

        private static int TryGetIntMember(object obj, string memberName)
        {
            if (obj == null || string.IsNullOrWhiteSpace(memberName))
            {
                return -1;
            }

            try
            {
                var value = obj.GetType().InvokeMember(
                    memberName,
                    System.Reflection.BindingFlags.GetProperty
                    | System.Reflection.BindingFlags.InvokeMethod
                    | System.Reflection.BindingFlags.Public
                    | System.Reflection.BindingFlags.Instance,
                    null,
                    obj,
                    Array.Empty<object>());
                return value is int intValue ? intValue : -1;
            }
            catch
            {
                return -1;
            }
        }

        private static string FormatValuePreview(object value)
        {
            if (value == null)
            {
                return "<null>";
            }

            if (value is string text)
            {
                return text.Length <= 160 ? text : text.Substring(0, 160) + "...";
            }

            if (value is string[] texts)
            {
                var joined = string.Join("; ", texts);
                return joined.Length <= 160 ? joined : joined.Substring(0, 160) + "...";
            }

            if (value is byte[] bytes)
            {
                return $"byte[{bytes.Length}]";
            }

            return value.ToString();
        }

        private static T SafeGet<T>(Func<T> getter, T fallback)
        {
            try
            {
                return getter();
            }
            catch
            {
                return fallback;
            }
        }

        private static void TryReleaseCom(object obj)
        {
            try
            {
                if (obj != null && Marshal.IsComObject(obj))
                {
                    Marshal.FinalReleaseComObject(obj);
                }
            }
            catch
            {
                // ignore
            }
        }

        private static object GetItem(object collection, object index)
        {
            if (collection == null)
            {
                return null;
            }

            try
            {
                return collection.GetType().InvokeMember(
                    "Item",
                    System.Reflection.BindingFlags.GetProperty
                    | System.Reflection.BindingFlags.InvokeMethod
                    | System.Reflection.BindingFlags.Public
                    | System.Reflection.BindingFlags.Instance,
                    null,
                    collection,
                    new[] { index });
            }
            catch
            {
                return null;
            }
        }

        internal static int Run(bool logAll, string tag)
        {
            MicroEngActions.Init();
            MicroEngActions.Log($"{tag}: start");

            try
            {
                var state = ComBridge.State;
                if (state == null)
                {
                    MicroEngActions.Log($"{tag}: ComApiBridge.State is null.");
                    MessageBox.Show(
                        "ComApiBridge.State is null. See MicroEng log for details.",
                        "MicroEng",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return 0;
                }

                var plugins = state.Plugins();
                if (plugins == null)
                {
                    MicroEngActions.Log($"{tag}: state.Plugins() returned null.");
                    return 0;
                }

                var count = SafeGet(() => plugins.Count, 0);
                MicroEngActions.Log($"{tag}: count={count}, logAll={logAll}");

                int matchCount = 0;
                for (int i = 1; i <= count; i++)
                {
                    object pluginObj = null;
                    try
                    {
                        pluginObj = GetItem(plugins, i);
                        if (pluginObj == null)
                        {
                            continue;
                        }

                        var plugin = pluginObj as ComApi.InwPlugin;
                        var typeName = pluginObj.GetType().FullName ?? "<unknown>";
                        var isSelectionTree = pluginObj is ComApi.InwSelectionTreePlugin;

                        var display = SafeGet(() => plugin?.iGetDisplayName(), null)
                            ?? TryGetStringMember(pluginObj, "iGetDisplayName");
                        var name = SafeGet(() => plugin?.ObjectName, null)
                            ?? TryGetStringMember(pluginObj, "ObjectName")
                            ?? TryGetStringMember(pluginObj, "name");

                        if (!logAll && !MatchesNeedles(display) && !MatchesNeedles(name) && !MatchesNeedles(typeName))
                        {
                            continue;
                        }

                        matchCount++;
                        var paramCount = SafeGet(() => plugin?.iGetNumParameters() ?? 0, -1);
                        if (paramCount < 0)
                        {
                            paramCount = TryGetIntMember(pluginObj, "iGetNumParameters");
                        }
                        if (paramCount < 0)
                        {
                            paramCount = 0;
                        }
                        MicroEngActions.Log($"{tag}: [{i}] name=\"{name}\" display=\"{display}\" type=\"{typeName}\" selectionTree={isSelectionTree} params={paramCount}");

                        if (paramCount > 0)
                        {
                            var zeroBased = TryGetParameter(plugin, pluginObj, 0, out _, out _);
                            var start = zeroBased ? 0 : 1;
                            var end = zeroBased ? paramCount - 1 : paramCount;

                            for (int p = start; p <= end; p++)
                            {
                                if (!TryGetParameter(plugin, pluginObj, p, out var paramName, out var paramValue))
                                {
                                    continue;
                                }

                                var preview = FormatValuePreview(paramValue);
                                MicroEngActions.Log($"{tag}:   param[{p}] {paramName} = {preview}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MicroEngActions.Log($"{tag}: item[{i}] failed ({ex.Message})");
                    }
                    finally
                    {
                        TryReleaseCom(pluginObj);
                    }
                }

                MicroEngActions.Log($"{tag}: matches={matchCount}");
                MessageBox.Show(
                    "COM plugin list probe complete. See MicroEng log for details.",
                    "MicroEng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"{tag} failed: {ex}");
                MessageBox.Show(
                    $"COM plugin list probe failed:\n{ex.Message}",
                    "MicroEng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            return 0;
        }
    }

    [Plugin("MicroEng.Dev.ComPluginListProbeAll", "MENG",
        DisplayName = "MicroEng Dev: COM Plugin List Probe (All)",
        ToolTip = "Logs all COM plugins to help identify Selection Tree-related plugins.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public sealed class DevComPluginListProbeAll : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            return DevComPluginListProbe.Run(logAll: true, tag: "ComPluginListProbeAll");
        }
    }
}
