using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using Microsoft.Win32;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

namespace MicroEng.Navisworks.SelectionTreeCom
{
    [Plugin("MicroEng.Dev.TreeMapperSelectionTree.RegisterCom", "MENG",
        DisplayName = "MicroEng Dev: COM Register TreeMapper SelectionTree",
        ToolTip = "Registers COM classes for the TreeMapper Selection Tree provider.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class DevRegisterTreeMapperSelectionTreeCom : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();
            MicroEngActions.Log("TreeMapper SelectionTree COM register: start");

            try
            {
                var asm = typeof(MicroEng.SelectionTreeCom.TreeMapperSelectionTreePlugin).Assembly;
                SelectionTreeComRegistrar.RegisterComPerUser(
                    asm,
                    typeof(MicroEng.SelectionTreeCom.TreeMapperSelectionTreePlugin),
                    "MicroEng.TreeMapperSelectionTreePlugin");

                var regInfo = SelectionTreeComRegistrar.RegisterNavisworksComPlugin("MicroEng.TreeMapperSelectionTreePlugin");
                MicroEngActions.Log($"TreeMapper SelectionTree COM register (per-user) ok, asm={asm.FullName}, loc={asm.Location}");
                MicroEngActions.Log(regInfo);
                MessageBox.Show(
                    $"COM registration completed (per-user).\n\nAssembly: {asm.Location}\n\n{regInfo}",
                    "MicroEng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree COM register failed: {ex}");
                MessageBox.Show(
                    $"COM registration failed:\n{ex.Message}",
                    "MicroEng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            return 0;
        }
    }

    [Plugin("MicroEng.Dev.TreeMapperSelectionTree.Install", "MENG",
        DisplayName = "MicroEng Dev: Install TreeMapper Selection Tree Option",
        ToolTip = "Writes a UserSelectionTree spec so TreeMapper appears in the Selection Tree dropdown.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class DevInstallTreeMapperSelectionTreeOption : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();
            MicroEngActions.Log("TreeMapper SelectionTree install: start");

            try
            {
                var state = ComBridge.State;
                if (state == null)
                {
                    throw new InvalidOperationException("ComApiBridge.State is null.");
                }

                var probeStartUtc = DateTime.UtcNow;
                LogAppSpecCandidates(state, "state");
                LogKnownNavisworksRoots();

                var editOpened = TryBeginEdit(state);
                try
                {
                    if (!TryCreateUserSelectionTreePlugin(state, out var pluginObj, out var createInfo))
                    {
                        MicroEngActions.Log($"TreeMapper SelectionTree install failed: {createInfo}");
                        MessageBox.Show(
                            $"Install failed:\n{createInfo}",
                            "MicroEng",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return 0;
                    }

                    LogAppSpecCandidates(pluginObj, "plugin");

                    ComApi.InwOpUserSelectionTreeSpec spec = null;
                    if (pluginObj is ComApi.InwOpUserSelectionTreePlugin pluginForSpec)
                    {
                        spec = TryCloneSpecFromPlugin(pluginForSpec);
                    }

                    if (spec == null)
                    {
                        spec = (ComApi.InwOpUserSelectionTreeSpec)state.ObjectFactory(
                            ComApi.nwEObjectType.eObjectType_nwOpUserSelectionTreeSpec, null, null);
                        MicroEngActions.Log("TreeMapper SelectionTree: created new UserSelectionTree spec.");
                    }

                    spec.UserString = "TreeMapper";
                    ConfigureSpecDefaults(spec);
                    TrySeedUserFindSpecs(state, spec);
                    TrySetProperty(spec, "nwOwn", true);
                    TrySetProperty(spec, "nwLock", 1);
                    LogOwnership(spec, "spec");
                    TryWriteSpecSnapshot(spec);

                    if (pluginObj is ComApi.InwOpUserSelectionTreePlugin plugin)
                    {
                        LogPluginIdentity(plugin, pluginObj);
                        LogOwnership(pluginObj, "plugin (before)");
                        var target = GetWritablePlugin(plugin);
                        if (!ReferenceEquals(target, plugin))
                        {
                            MicroEngActions.Log("TreeMapper SelectionTree: using copy() for writable plugin instance.");
                        }

                        TrySetProperty(target, "nwOwn", true);
                        TrySetProperty(target, "nwLock", 1);
                        TrySetPluginProperty(target, "name", "MicroEng.TreeMapperSelectionTreePlugin");
                        TrySetPluginProperty(target, "Version", 1);
                        target.SetTreeSpec(spec);
                        LogOwnership(target, "plugin (after set own/lock)");
                        LogTreeSpec(target);
                        if (!TrySaveSpec(target, pluginObj, spec, out var saveMessage))
                        {
                            throw new InvalidOperationException(saveMessage);
                        }
                        MicroEngActions.Log($"TreeMapper SelectionTree: {saveMessage}");
                        TryRegisterUserTreeSpecEntries();
                    }
                    else
                    {
                        LogPluginIdentity(null, pluginObj);
                        LogOwnership(pluginObj, "plugin (before)");
                        TrySetProperty(pluginObj, "nwOwn", true);
                        TrySetProperty(pluginObj, "nwLock", 1);
                        TryInvoke(pluginObj, "SetTreeSpec", spec);
                        LogOwnership(pluginObj, "plugin (after set own/lock)");
                        if (!TrySaveSpec(null, pluginObj, spec, out var saveMessage))
                        {
                            throw new InvalidOperationException(saveMessage);
                        }
                        MicroEngActions.Log($"TreeMapper SelectionTree: {saveMessage}");
                        TryRegisterUserTreeSpecEntries();
                    }

                    MicroEngActions.Log("TreeMapper SelectionTree spec saved to app spec dir.");
                    MessageBox.Show(
                        "Installed TreeMapper Selection Tree option.\n\nRestart Navisworks, then open Selection Tree dropdown and look for \"TreeMapper\".",
                        "MicroEng",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                finally
                {
                    if (editOpened)
                    {
                        TryEndEdit(state);
                    }

                    LogRecentSpecWrites(probeStartUtc);
                }

            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree install failed: {ex}");
                MessageBox.Show(
                    $"Install failed:\n{ex}",
                    "MicroEng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            return 0;
        }

        private static bool TryCreateUserSelectionTreePlugin(ComApi.InwOpState state, out object pluginObj, out string info)
        {
            pluginObj = null;
            info = "";

            if (TryCreateUserSelectionTreePluginViaCreatePlugin(state, out pluginObj, out info))
            {
                return true;
            }

            if (TryGetUserSelectionTreePluginFromPlugins(state, out pluginObj, out info))
            {
                return true;
            }

            if (TryCreateViaObjectFactory(state, out pluginObj, out info))
            {
                return true;
            }

            string[] progIds =
            {
                "nwOpUserSelectionTreePlugin",
                "InwOpUserSelectionTreePlugin",
                "Autodesk.Navisworks.Api.Interop.ComApi.InwOpUserSelectionTreePlugin"
            };

            foreach (var progId in progIds)
            {
                try
                {
                    var type = Type.GetTypeFromProgID(progId, throwOnError: false);
                    if (type == null)
                    {
                        continue;
                    }

                    pluginObj = Activator.CreateInstance(type);
                    if (pluginObj != null)
                    {
                        info = $"Created COM object via ProgID '{progId}'.";
                        MicroEngActions.Log(info);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    MicroEngActions.Log($"ProgID create failed ({progId}): {ex.Message}");
                }
            }

            try
            {
                var clsid = new Guid("B98050D4-8318-4FD5-9590-5B62A151180F");
                pluginObj = Activator.CreateInstance(Type.GetTypeFromCLSID(clsid, throwOnError: false));
                if (pluginObj != null)
                {
                    info = $"Created COM object via CLSID {clsid}.";
                    MicroEngActions.Log(info);
                    return true;
                }
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"CLSID create failed: {ex.Message}");
            }

            info = "Could not create InwOpUserSelectionTreePlugin COM object (ObjectFactory scan + ProgID/CLSID failed).";
            return false;
        }

        private static bool TryGetUserSelectionTreePluginFromPlugins(ComApi.InwOpState state, out object pluginObj, out string info)
        {
            pluginObj = null;
            info = "";

            if (state == null)
            {
                info = "ComApi state is null.";
                return false;
            }

            try
            {
                var plugins = state.Plugins();
                if (plugins == null)
                {
                    info = "state.Plugins() returned null.";
                    return false;
                }

                var count = plugins.Count;
                for (int i = 1; i <= count; i++)
                {
                    object obj = null;
                    try
                    {
                        obj = GetPluginsItem(plugins, i);
                        if (obj == null)
                        {
                            continue;
                        }

                        var name = GetPluginName(obj);
                        if (string.Equals(name, "nwOpUserSelectionTreePlugin", StringComparison.OrdinalIgnoreCase)
                            || LooksLikeSelectionTreePlugin(obj))
                        {
                            pluginObj = obj;
                            info = $"Found plugin in state.Plugins() at index {i} (name={name ?? "<unknown>"}).";
                            MicroEngActions.Log(info);
                            return true;
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                    finally
                    {
                        if (!ReferenceEquals(obj, pluginObj))
                        {
                            TryReleaseCom(obj);
                        }
                    }
                }

                info = "state.Plugins() did not contain a UserSelectionTree plugin.";
                return false;
            }
            catch (Exception ex)
            {
                info = $"state.Plugins() scan failed: {ex.Message}";
                return false;
            }
        }

        private static bool TryCreateUserSelectionTreePluginViaCreatePlugin(ComApi.InwOpState state, out object pluginObj, out string info)
        {
            pluginObj = null;
            info = "";

            if (state == null)
            {
                info = "ComApi state is null.";
                return false;
            }

            string[] names =
            {
                "nwOpUserSelectionTreePlugin",
                "UserSelectionTreePlugin",
                "InwOpUserSelectionTreePlugin"
            };

            foreach (var name in names)
            {
                try
                {
                    var result = state.CreatePlugin(name);
                    MicroEngActions.Log($"CreatePlugin({name}) returned {result}.");
                    if (result <= 0)
                    {
                        continue;
                    }

                    var plugins = state.Plugins();
                    pluginObj = GetPluginsItem(plugins, result);
                    if (pluginObj != null)
                    {
                        info = $"Created plugin via CreatePlugin(\"{name}\") at index {result}.";
                        MicroEngActions.Log(info);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    MicroEngActions.Log($"CreatePlugin({name}) failed: {ex.Message}");
                }
            }

            info = "CreatePlugin attempts did not return a usable UserSelectionTree plugin.";
            return false;
        }

        private static void TrySetPluginProperty(object pluginObj, string propertyName, object value)
        {
            try
            {
                var prop = pluginObj.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
                if (prop == null || !prop.CanWrite)
                {
                    MicroEngActions.Log($"TreeMapper SelectionTree: set {propertyName} skipped (read-only or missing).");
                    return;
                }

                pluginObj.GetType().InvokeMember(
                    propertyName,
                    System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public,
                    null,
                    pluginObj,
                    new[] { value });
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: set {propertyName} failed ({ex.Message})");
            }
        }

        private static void TrySetProperty(object target, string propertyName, object value)
        {
            if (target == null)
            {
                return;
            }

            try
            {
                target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    target,
                    new[] { value });
            }
            catch
            {
                // ignore
            }
        }

        private static void TryInvoke(object pluginObj, string methodName, params object[] args)
        {
            try
            {
                pluginObj.GetType().InvokeMember(
                    methodName,
                    System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public,
                    null,
                    pluginObj,
                    args);
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: call {methodName} failed ({ex.Message})");
            }
        }

        private static bool TryBeginEdit(ComApi.InwOpState state)
        {
            if (state == null)
            {
                return false;
            }

            try
            {
                state.BeginEdit("MicroEng.TreeMapperSelectionTree");
                MicroEngActions.Log("TreeMapper SelectionTree: BeginEdit ok.");
                return true;
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: BeginEdit failed ({ex.Message}).");
                return false;
            }
        }

        private static void TryEndEdit(ComApi.InwOpState state)
        {
            if (state == null)
            {
                return;
            }

            try
            {
                state.EndEdit();
                MicroEngActions.Log("TreeMapper SelectionTree: EndEdit ok.");
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: EndEdit failed ({ex.Message}).");
            }
        }

        private static void LogOwnership(object obj, string label)
        {
            if (obj == null)
            {
                return;
            }

            var ro = TryGetBoolProperty(obj, "nwReadOnly");
            var own = TryGetBoolProperty(obj, "nwOwn");
            var lockValue = TryGetIntProperty(obj, "nwLock");
            MicroEngActions.Log($"TreeMapper SelectionTree: {label} nwReadOnly={ro}, nwOwn={own}, nwLock={lockValue}");
        }

        private static bool? TryGetBoolProperty(object obj, string propertyName)
        {
            try
            {
                var value = obj.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    obj,
                    Array.Empty<object>());
                return value is bool boolValue ? boolValue : (bool?)null;
            }
            catch
            {
                return null;
            }
        }

        private static int? TryGetIntProperty(object obj, string propertyName)
        {
            try
            {
                var value = obj.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    obj,
                    Array.Empty<object>());
                return value is int intValue ? intValue : (int?)null;
            }
            catch
            {
                return null;
            }
        }

        private static void TrySeedUserFindSpecs(ComApi.InwOpState state, ComApi.InwOpUserSelectionTreeSpec spec)
        {
            if (state == null || spec == null)
            {
                return;
            }

            try
            {
                var userFind = state.ObjectFactory(
                    ComApi.nwEObjectType.eObjectType_nwOpUserFindSpec,
                    null,
                    null) as ComApi.InwOpUserFindSpec;

                if (userFind == null)
                {
                    return;
                }

                TrySetProperty(userFind, "ExplicitName", "TreeMapper");
                TrySetProperty(userFind, "NameMode", ComApi.nwESelTreeNameMode.eSelTreeNameMode_EXPLICIT);
                TrySetProperty(userFind, "GroupMode", ComApi.nwESelTreeGroupMode.eSelTreeGroupMode_NONE);
                TrySetProperty(userFind, "TextFormatMode", ComApi.nwESelTreeTextFormatMode.nwESelTreeTextFormatMode_EXPLICIT);
                TrySetProperty(userFind, "ExplicitTextFormat", ComApi.nwESelTreeTextFormat.eSelTreeTxtFmt_NORMAL);
                TrySetProperty(userFind, "IconMode", ComApi.nwESelTreeIconMode.nwESelTreeIconMode_EXPLICIT);
                TrySetProperty(userFind, "ExplicitIcon", ComApi.nwESelTreeIcon.nwESelTreeIcon_SELECTION_SET_NORMAL);
                TrySetProperty(userFind, "SearchMode", ComApi.nwESearchMode.eSearchMode_ALL_PATHS);
                TrySetProperty(userFind, "SelectionSubPaths", true);

                TryAddPlaceholderCondition(state, userFind);

                var coll = spec.UserFindSpecs();
                if (coll == null)
                {
                    return;
                }

                coll.Add(userFind);
                MicroEngActions.Log("TreeMapper SelectionTree: seeded UserFindSpecs with explicit name.");
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: seed UserFindSpecs failed ({ex.Message})");
            }
        }

        private static void TryAddPlaceholderCondition(ComApi.InwOpState state, ComApi.InwOpUserFindSpec userFind)
        {
            if (state == null || userFind == null)
            {
                return;
            }

            try
            {
                var condition = state.ObjectFactory(
                    ComApi.nwEObjectType.eObjectType_nwOpFindCondition,
                    null,
                    null) as ComApi.InwOpFindCondition;

                if (condition == null)
                {
                    return;
                }

                // Minimal placeholder: Category "Item", Property "Name", condition "Has Property".
                condition.SetAttributeNames("Item", "Item");
                condition.SetPropertyNames("Name", "Name");
                condition.Condition = ComApi.nwEFindCondition.eFind_HAS_PROP;
                TrySetProperty(condition, "value", "");

                var conditions = userFind.Conditions();
                if (conditions == null)
                {
                    return;
                }

                conditions.Add(condition);
                MicroEngActions.Log("TreeMapper SelectionTree: added placeholder Find Condition (Item.Name exists).");
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: add placeholder condition failed ({ex.Message})");
            }
        }

        private static ComApi.InwOpUserSelectionTreeSpec TryCloneSpecFromPlugin(ComApi.InwOpUserSelectionTreePlugin plugin)
        {
            if (plugin == null)
            {
                return null;
            }

            try
            {
                var specObj = plugin.GetTreeSpec();
                if (specObj is ComApi.InwOpUserSelectionTreeSpec spec)
                {
                    try
                    {
                        var copyObj = spec.Copy();
                        if (copyObj is ComApi.InwOpUserSelectionTreeSpec copy)
                        {
                            MicroEngActions.Log("TreeMapper SelectionTree: cloned spec from plugin.");
                            return copy;
                        }
                    }
                    catch (Exception ex)
                    {
                        MicroEngActions.Log($"TreeMapper SelectionTree: spec copy failed ({ex.Message}).");
                    }

                    MicroEngActions.Log("TreeMapper SelectionTree: using plugin spec without copy.");
                    return spec;
                }
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: GetTreeSpec failed ({ex.Message}).");
            }

            return null;
        }

        private static bool TrySaveSpec(
            ComApi.InwOpUserSelectionTreePlugin plugin,
            object pluginObj,
            ComApi.InwOpUserSelectionTreeSpec spec,
            out string message)
        {
            message = "SaveFileToAppSpecDir succeeded.";

            if (spec == null)
            {
                message = "SaveFileToAppSpecDir failed: spec is null.";
                return false;
            }

            Exception saveError = null;
            try
            {
                if (plugin != null)
                {
                    plugin.SaveFileToAppSpecDir();
                    TryWriteSpecAliases(spec, plugin?.name ?? GetPluginName(pluginObj));
                    return true;
                }

                if (TryInvokeSaveFileToAppSpecDir(pluginObj, out saveError))
                {
                    TryWriteSpecAliases(spec, plugin?.name ?? GetPluginName(pluginObj));
                    return true;
                }
            }
            catch (Exception ex)
            {
                saveError = ex;
            }

            var pluginName = plugin?.name ?? GetPluginName(pluginObj) ?? "LcUntitledPlugin";
            if (TryWriteSpecToTreeSpecDir(spec, pluginName, out var fallbackPath, out var fallbackError))
            {
                TryWriteSpecAliases(spec, pluginName);
                message = $"SaveFileToAppSpecDir failed; wrote spec to '{fallbackPath}'.";
                return true;
            }

            var errorText = saveError?.Message ?? "unknown error";
            message = $"SaveFileToAppSpecDir failed ({errorText}). Fallback failed ({fallbackError}).";
            return false;
        }

        private static bool TryInvokeSaveFileToAppSpecDir(object pluginObj, out Exception exception)
        {
            exception = null;
            if (pluginObj == null)
            {
                return false;
            }

            try
            {
                pluginObj.GetType().InvokeMember(
                    "SaveFileToAppSpecDir",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    pluginObj,
                    Array.Empty<object>());
                return true;
            }
            catch (Exception ex)
            {
                exception = ex;
                return false;
            }
        }

        private static bool TryWriteSpecToTreeSpecDir(
            ComApi.InwOpUserSelectionTreeSpec spec,
            string pluginName,
            out string path,
            out string error)
        {
            path = null;
            error = null;

            try
            {
                var treeSpecDir = GetUserTreeSpecDir();
                var safeName = SanitizeFileName(string.IsNullOrWhiteSpace(pluginName) ? "LcUntitledPlugin" : pluginName);
                path = Path.Combine(treeSpecDir, $"{safeName}.spc");

                spec.SaveToFile(path);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "LcUntitledPlugin";
            }

            var invalid = Path.GetInvalidFileNameChars();
            var chars = name.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (Array.IndexOf(invalid, chars[i]) >= 0)
                {
                    chars[i] = '_';
                }
            }

            return new string(chars);
        }

        private static void TryWriteSpecAliases(ComApi.InwOpUserSelectionTreeSpec spec, string pluginName)
        {
            if (spec == null)
            {
                return;
            }

            try
            {
                var treeSpecDir = GetUserTreeSpecDir();

                var aliases = new[]
                {
                    "TreeMapper",
                    "MicroEng.TreeMapperSelectionTreePlugin"
                };

                foreach (var alias in aliases)
                {
                    var safeAlias = SanitizeFileName(alias);
                    var aliasPath = Path.Combine(treeSpecDir, $"{safeAlias}.spc");
                    if (string.Equals(aliasPath, pluginName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    try
                    {
                        spec.SaveToFile(aliasPath);
                        MicroEngActions.Log($"TreeMapper SelectionTree: wrote alias spec '{aliasPath}'.");
                    }
                    catch (Exception ex)
                    {
                        MicroEngActions.Log($"TreeMapper SelectionTree: alias write failed ({aliasPath}) ({ex.Message}).");
                    }
                }
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: alias write failed ({ex.Message}).");
            }
        }

        private static void TryRegisterUserTreeSpecEntries()
        {
            try
            {
                var treeSpecDir = GetUserTreeSpecDir();

                var entries = new[]
                {
                    ("TreeMapper", Path.Combine(treeSpecDir, "TreeMapper.spc")),
                    ("MicroEng.TreeMapperSelectionTreePlugin", Path.Combine(treeSpecDir, "MicroEng.TreeMapperSelectionTreePlugin.spc")),
                    ("LcUntitledPlugin", Path.Combine(treeSpecDir, "LcUntitledPlugin.spc"))
                };

                WriteUserTreeSpecRegistryEntries(Registry.CurrentUser, "HKCU", entries);
                WriteUserTreeSpecRegistryEntries(Registry.LocalMachine, "HKLM", entries);
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: registry write failed ({ex.Message}).");
            }
        }

        private static void WriteUserTreeSpecRegistryEntries(
            RegistryKey root,
            string rootName,
            (string name, string path)[] entries)
        {
            try
            {
                const string basePath = @"Software\Autodesk\Navisworks Manage";
                using var baseKey = root.CreateSubKey(basePath);
                if (baseKey == null)
                {
                    MicroEngActions.Log($"TreeMapper SelectionTree: {rootName} base registry path unavailable.");
                    return;
                }

                var versions = GetVersionSubKeys(baseKey);
                if (versions == null || versions.Count == 0)
                {
                    MicroEngActions.Log($"TreeMapper SelectionTree: {rootName} Navisworks version key not found.");
                    return;
                }

                foreach (var version in versions)
                {
                    var specsPath = $@"{basePath}\{version}\User Tree Specs";
                    using var specKey = root.CreateSubKey(specsPath);
                    if (specKey == null)
                    {
                        MicroEngActions.Log($"TreeMapper SelectionTree: {rootName} failed to open {specsPath}.");
                        continue;
                    }

                    foreach (var entry in entries)
                    {
                        try
                        {
                            specKey.SetValue(entry.name, entry.path, RegistryValueKind.String);
                            MicroEngActions.Log($"TreeMapper SelectionTree: {rootName} User Tree Specs [{version}] set {entry.name}={entry.path}");
                        }
                        catch (Exception ex)
                        {
                            MicroEngActions.Log($"TreeMapper SelectionTree: {rootName} User Tree Specs [{version}] write failed ({entry.name}) ({ex.Message}).");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: {rootName} User Tree Specs write failed ({ex.Message}).");
            }
        }

        private static List<string> GetVersionSubKeys(RegistryKey baseKey)
        {
            try
            {
                var candidates = baseKey.GetSubKeyNames();
                if (candidates == null || candidates.Length == 0)
                {
                    return null;
                }

                var versions = new List<(Version version, string name)>();
                foreach (var name in candidates)
                {
                    if (Version.TryParse(name, out var version))
                    {
                        versions.Add((version, name));
                    }
                }

                if (versions.Count == 0)
                {
                    return null;
                }

                versions.Sort((a, b) => b.version.CompareTo(a.version));
                return versions.Select(v => v.name).ToList();
            }
            catch
            {
                return null;
            }
        }

        private static string GetUserTreeSpecDir()
        {
            var treeSpecDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Autodesk",
                "Navisworks Manage 2025",
                "TreeSpec");
            Directory.CreateDirectory(treeSpecDir);
            return treeSpecDir;
        }

        private static void LogAppSpecCandidates(object obj, string label)
        {
            if (obj == null)
            {
                return;
            }

            string[] candidates =
            {
                "AppSpecDir",
                "AppSpecPath",
                "AppSpec",
                "SpecDir",
                "SpecPath",
                "UserSpecDir",
                "UserSpecPath",
                "SelectionTreeSpecDir",
                "SelectionTreeSpecPath",
                "AppDataDir",
                "RoamingDir",
                "AppDataCurrentDirectory",
                "AppDataInheritedDirectory",
                "AppDataSharedDirectory",
                "GetAppSpecDir",
                "GetAppSpecFileExt"
            };

            foreach (var name in candidates)
            {
                var value = TryGetStringMember(obj, name);
                if (string.IsNullOrWhiteSpace(value))
                {
                    value = TryGetXtensionString(obj, name);
                }

                if (string.IsNullOrWhiteSpace(value))
                {
                    var raw = TryGetXtensionValue(obj, name);
                    if (raw != null)
                    {
                        value = CoerceString(raw);
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            value = raw.GetType().FullName;
                        }
                    }
                }

                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                var exists = Directory.Exists(value) ? "exists" : "missing";
                MicroEngActions.Log($"TreeMapper SelectionTree: {label} {name}={value} ({exists})");
            }
        }

        private static void LogKnownNavisworksRoots()
        {
            var roots = GetNavisworksRoots();
            foreach (var root in roots)
            {
                var exists = Directory.Exists(root) ? "exists" : "missing";
                MicroEngActions.Log($"TreeMapper SelectionTree: root {root} ({exists})");
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
                    BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
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

        private static string TryGetXtensionString(object obj, string key)
        {
            if (obj == null || string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            try
            {
                var value = obj.GetType().InvokeMember(
                    "Xtension",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    obj,
                    new object[] { key });
                return CoerceString(value);
            }
            catch
            {
                return null;
            }
        }

        private static object TryGetXtensionValue(object obj, string key)
        {
            if (obj == null || string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            try
            {
                return obj.GetType().InvokeMember(
                    "Xtension",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    obj,
                    new object[] { key });
            }
            catch
            {
                return null;
            }
        }

        private static string CoerceString(object value)
        {
            if (value == null)
            {
                return null;
            }

            if (value is string text)
            {
                return text;
            }

            try
            {
                return value.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static void LogRecentSpecWrites(DateTime sinceUtc)
        {
            try
            {
                var roots = GetNavisworksRoots();
                var results = new System.Collections.Generic.List<string>();

                foreach (var root in roots)
                {
                    if (!Directory.Exists(root))
                    {
                        continue;
                    }

                    foreach (var file in Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories))
                    {
                        DateTime writeTimeUtc;
                        try
                        {
                            writeTimeUtc = File.GetLastWriteTimeUtc(file);
                        }
                        catch
                        {
                            continue;
                        }

                        if (writeTimeUtc >= sinceUtc)
                        {
                            results.Add($"{file} ({writeTimeUtc:O})");
                            if (results.Count >= 50)
                            {
                                break;
                            }
                        }
                    }

                    if (results.Count >= 50)
                    {
                        break;
                    }
                }

                MicroEngActions.Log($"TreeMapper SelectionTree: recent file writes since {sinceUtc:O} = {results.Count}");
                foreach (var entry in results)
                {
                    MicroEngActions.Log($"TreeMapper SelectionTree:   {entry}");
                }
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: recent file write scan failed ({ex.Message}).");
            }
        }

        private static string[] GetNavisworksRoots()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);

            return new[]
            {
                Path.Combine(appData, "Autodesk", "Navisworks Manage 2025"),
                Path.Combine(localAppData, "Autodesk", "Navisworks 2025"),
                Path.Combine(programData, "Autodesk", "Navisworks Manage 2025"),
                Path.Combine(programFiles, "Autodesk", "Navisworks Manage 2025")
            };
        }

        private static void LogPluginIdentity(ComApi.InwOpUserSelectionTreePlugin plugin, object pluginObj)
        {
            try
            {
                var typeName = pluginObj?.GetType().FullName ?? "<null>";
                var name = plugin?.name ?? GetPluginName(pluginObj);
                var version = plugin?.Version ?? TryGetPluginVersion(pluginObj);
                MicroEngActions.Log($"TreeMapper SelectionTree: plugin type={typeName}, name={name ?? "<null>"}, version={version}");
            }
            catch
            {
                // ignore
            }
        }

        private static int TryGetPluginVersion(object pluginObj)
        {
            if (pluginObj == null)
            {
                return 0;
            }

            try
            {
                var value = pluginObj.GetType().InvokeMember(
                    "Version",
                    BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    pluginObj,
                    Array.Empty<object>());
                return value is int intValue ? intValue : 0;
            }
            catch
            {
                return 0;
            }
        }

        private static void ConfigureSpecDefaults(ComApi.InwOpUserSelectionTreeSpec spec)
        {
            if (spec == null)
            {
                return;
            }

            try
            {
                spec.Top = ComApi.nwESelTreeTop.eSelTreeTop_MODEL;
                spec.Bottom = ComApi.nwESelTreeBottom.eSelTreeBottom_GEOMETRY;
                MicroEngActions.Log("TreeMapper SelectionTree: spec defaults set (Top=MODEL, Bottom=GEOMETRY).");
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: set spec defaults failed ({ex.Message})");
            }
        }

        private static ComApi.InwOpUserSelectionTreePlugin GetWritablePlugin(ComApi.InwOpUserSelectionTreePlugin plugin)
        {
            if (plugin == null)
            {
                return null;
            }

            try
            {
                var readOnly = plugin.nwReadOnly;
                MicroEngActions.Log($"TreeMapper SelectionTree: plugin readOnly={readOnly}, own={plugin.nwOwn}, lock={plugin.nwLock}");
                if (!readOnly)
                {
                    return plugin;
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var copyObj = plugin.Copy();
                if (copyObj is ComApi.InwOpUserSelectionTreePlugin copy)
                {
                    return copy;
                }
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: plugin copy failed ({ex.Message})");
            }

            return plugin;
        }

        private static void LogTreeSpec(ComApi.InwOpUserSelectionTreePlugin plugin)
        {
            if (plugin == null)
            {
                return;
            }

            try
            {
                var specObj = plugin.GetTreeSpec();
                if (specObj is ComApi.InwOpUserSelectionTreeSpec spec)
                {
                    MicroEngActions.Log($"TreeMapper SelectionTree: spec UserString={spec.UserString}, Top={spec.Top}, Bottom={spec.Bottom}");
                }
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree: GetTreeSpec failed ({ex.Message})");
            }
        }

        private static bool TryCreateViaObjectFactory(ComApi.InwOpState state, out object pluginObj, out string info)
        {
            pluginObj = null;
            info = "";

            if (state == null)
            {
                info = "ComApi state is null.";
                return false;
            }

            const int minTypeId = 1;
            const int maxTypeId = 2000;
            var sw = Stopwatch.StartNew();

            for (int typeId = minTypeId; typeId <= maxTypeId; typeId++)
            {
                if (sw.ElapsedMilliseconds > 3000)
                {
                    info = $"ObjectFactory scan timed out after {sw.ElapsedMilliseconds}ms.";
                    return false;
                }

                object obj;
                try
                {
                    obj = state.GetType().InvokeMember(
                        "ObjectFactory",
                        BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                        null,
                        state,
                        new object[] { typeId, null, null });
                }
                catch
                {
                    continue;
                }

                if (obj == null)
                {
                    continue;
                }

                if (LooksLikeSelectionTreePlugin(obj))
                {
                    pluginObj = obj;
                    info = $"Created plugin via ObjectFactory(typeId={typeId}).";
                    MicroEngActions.Log(info);
                    return true;
                }

                TryReleaseCom(obj);
            }

            info = $"ObjectFactory scan {minTypeId}-{maxTypeId} did not find a UserSelectionTree plugin.";
            return false;
        }

        private static bool LooksLikeSelectionTreePlugin(object obj)
        {
            var type = obj.GetType();
            return type.GetMethod("SetTreeSpec") != null
                && type.GetMethod("SaveFileToAppSpecDir") != null;
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

        private static void TryWriteSpecSnapshot(ComApi.InwOpUserSelectionTreeSpec spec)
        {
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MicroEng.Navisworks",
                    "Diagnostics");
                Directory.CreateDirectory(dir);

                var path = Path.Combine(dir, $"TreeMapperSelectionTreeSpec_{DateTime.Now:yyyyMMdd_HHmmss}.dat");
                spec.SaveToFile(path);

                MicroEngActions.Log($"TreeMapper SelectionTree spec snapshot saved: {path}");
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper SelectionTree spec snapshot failed: {ex.Message}");
            }
        }

        private static object GetPluginsItem(object plugins, object index)
        {
            if (plugins == null)
            {
                return null;
            }

            try
            {
                return plugins.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    plugins,
                    new[] { index });
            }
            catch
            {
                return null;
            }
        }

        private static string GetPluginName(object pluginObj)
        {
            if (pluginObj == null)
            {
                return null;
            }

            try
            {
                return pluginObj.GetType().InvokeMember(
                    "ObjectName",
                    BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    pluginObj,
                    Array.Empty<object>()) as string;
            }
            catch
            {
                // ignore
            }

            try
            {
                return pluginObj.GetType().InvokeMember(
                    "name",
                    BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    pluginObj,
                    Array.Empty<object>()) as string;
            }
            catch
            {
                return null;
            }
        }

    }

    internal static class SelectionTreeComRegistrar
    {
        public static void RegisterComPerUser(Assembly assembly, Type comType, string progId)
        {
            if (assembly == null) throw new ArgumentNullException(nameof(assembly));
            if (comType == null) throw new ArgumentNullException(nameof(comType));

            var clsid = comType.GUID;
            if (clsid == Guid.Empty)
            {
                throw new InvalidOperationException("COM type GUID is empty.");
            }

            var clsidPath = $@"Software\Classes\CLSID\{{{clsid}}}";
            using (var clsidKey = Registry.CurrentUser.CreateSubKey(clsidPath))
            {
                if (clsidKey == null)
                {
                    throw new InvalidOperationException("Failed to create CLSID registry key.");
                }

                clsidKey.SetValue(null, progId);
                clsidKey.CreateSubKey("ProgId")?.SetValue(null, progId);
                clsidKey.CreateSubKey("Implemented Categories\\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}");

                using (var inproc = clsidKey.CreateSubKey("InprocServer32"))
                {
                    if (inproc == null)
                    {
                        throw new InvalidOperationException("Failed to create InprocServer32 registry key.");
                    }

                    inproc.SetValue(null, "mscoree.dll");
                    inproc.SetValue("ThreadingModel", "Both");
                    inproc.SetValue("Class", comType.FullName ?? comType.Name);
                    inproc.SetValue("Assembly", assembly.FullName ?? "");
                    inproc.SetValue("RuntimeVersion", assembly.ImageRuntimeVersion ?? "");
                    inproc.SetValue("CodeBase", assembly.Location);
                }
            }

            using (var progIdKey = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{progId}"))
            {
                progIdKey?.SetValue(null, progId);
                progIdKey?.CreateSubKey("CLSID")?.SetValue(null, $"{{{clsid}}}");
            }
        }

        public static string RegisterNavisworksComPlugin(string progId)
        {
            if (string.IsNullOrWhiteSpace(progId))
            {
                return "Navisworks COM plugin registration skipped (ProgID is empty).";
            }

            var results = new[]
            {
                TryRegisterNavisworksComPlugin(Registry.CurrentUser, progId, "HKCU"),
                TryRegisterNavisworksComPlugin(Registry.LocalMachine, progId, "HKLM")
            };

            return string.Join(" | ", results);
        }

        private static string TryRegisterNavisworksComPlugin(RegistryKey root, string progId, string rootName)
        {
            try
            {
                const string basePath = @"Software\Autodesk\Navisworks Manage";
                using var baseKey = root.CreateSubKey(basePath);
                if (baseKey == null)
                {
                    return $"{rootName}: Navisworks COM Plugins key not available.";
                }

                var versions = GetVersionSubKeys(baseKey);
                if (versions == null || versions.Count == 0)
                {
                    return $"{rootName}: Navisworks version key not found under {basePath}.";
                }

                var results = new List<string>();
                foreach (var version in versions)
                {
                    var comPluginsPath = $@"{basePath}\{version}\COM Plugins";
                    using var comKey = root.CreateSubKey(comPluginsPath);
                    if (comKey == null)
                    {
                        results.Add($"{rootName} [{version}]: Failed to open {comPluginsPath}.");
                        continue;
                    }

                    comKey.SetValue(progId, string.Empty, RegistryValueKind.String);

                    var licensePath = $@"{basePath}\{version}\COM Plugin License";
                    using var licenseKey = root.CreateSubKey(licensePath);
                    if (licenseKey == null)
                    {
                        results.Add($"{rootName} [{version}]: Registered COM plugin '{progId}' under {comPluginsPath} (license key missing).");
                        continue;
                    }

                    licenseKey.SetValue(progId, 1, RegistryValueKind.DWord);
                    results.Add($"{rootName} [{version}]: Registered COM plugin '{progId}' under {comPluginsPath} (license=1).");
                }

                return string.Join(" | ", results);
            }
            catch (UnauthorizedAccessException ex)
            {
                return $"{rootName}: Registry write denied ({ex.Message}).";
            }
            catch (Exception ex)
            {
                return $"{rootName}: Registry write failed ({ex.Message}).";
            }
        }

        private static List<string> GetVersionSubKeys(RegistryKey baseKey)
        {
            try
            {
                var candidates = baseKey.GetSubKeyNames();
                if (candidates == null || candidates.Length == 0)
                {
                    return null;
                }

                var versions = new List<(Version version, string name)>();
                foreach (var name in candidates)
                {
                    if (Version.TryParse(name, out var version))
                    {
                        versions.Add((version, name));
                    }
                }

                if (versions.Count == 0)
                {
                    return null;
                }

                versions.Sort((a, b) => b.version.CompareTo(a.version));
                return versions.Select(v => v.name).ToList();
            }
            catch
            {
                return null;
            }
        }

        private static string GetLatestVersionSubKey(RegistryKey baseKey)
        {
            try
            {
                var candidates = baseKey.GetSubKeyNames();
                if (candidates == null || candidates.Length == 0)
                {
                    return null;
                }

                Version bestVersion = null;
                string bestName = null;
                foreach (var name in candidates)
                {
                    if (!Version.TryParse(name, out var version))
                    {
                        continue;
                    }

                    if (bestVersion == null || version > bestVersion)
                    {
                        bestVersion = version;
                        bestName = name;
                    }
                }

                return bestName ?? (candidates.Length > 0 ? candidates[0] : null);
            }
            catch
            {
                return null;
            }
        }
    }
}
