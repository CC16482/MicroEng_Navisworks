using System;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using Microsoft.Win32;

namespace MicroEng.Navisworks.SelectionTreeCom
{
    [Plugin("MicroEng.Dev.SelectionTree.RegistryProbe", "MENG",
        DisplayName = "MicroEng Dev: Selection Tree Registry Probe",
        ToolTip = "Dumps Navisworks Selection Tree-related registry keys to the MicroEng log.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public sealed class DevSelectionTreeRegistryProbe : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();
            MicroEngActions.Log("SelectionTreeRegistryProbe: start");

            try
            {
                var root = GetNavisworksRegistryRoot();
                if (string.IsNullOrWhiteSpace(root))
                {
                    MicroEngActions.Log("SelectionTreeRegistryProbe: Navisworks registry root not found under HKCU\\Software\\Autodesk\\Navisworks Manage.");
                    MessageBox.Show(
                        "Navisworks registry root not found. See MicroEng log for details.",
                        "MicroEng",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return 0;
                }

                MicroEngActions.Log($"SelectionTreeRegistryProbe: root={root}");

                LogKeySummary(root, "Selection Tree");
                LogKeySummary(root, "User Tree Specs");
                LogKeySummary(root, @"GlobalOptions\interface\selection_tree");

                MessageBox.Show(
                    "Selection Tree registry probe complete. See MicroEng log for details.",
                    "MicroEng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"SelectionTreeRegistryProbe failed: {ex}");
                MessageBox.Show(
                    $"Selection Tree registry probe failed:\n{ex.Message}",
                    "MicroEng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            return 0;
        }

        private static string GetNavisworksRegistryRoot()
        {
            const string basePath = @"Software\Autodesk\Navisworks Manage";
            using var baseKey = Registry.CurrentUser.OpenSubKey(basePath);
            if (baseKey == null)
            {
                return null;
            }

            var candidates = baseKey.GetSubKeyNames();
            if (candidates == null || candidates.Length == 0)
            {
                return null;
            }

            var best = candidates
                .Select(name => new { Name = name, Version = TryParseVersion(name) })
                .Where(x => x.Version != null)
                .OrderByDescending(x => x.Version)
                .FirstOrDefault();

            var selected = best?.Name ?? candidates.FirstOrDefault();
            return string.IsNullOrWhiteSpace(selected) ? null : $@"{basePath}\{selected}";
        }

        private static Version TryParseVersion(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return null;
            }

            return Version.TryParse(input, out var version) ? version : null;
        }

        private static void LogKeySummary(string root, string subKey)
        {
            var path = string.IsNullOrWhiteSpace(subKey) ? root : $@"{root}\{subKey}";
            using var key = Registry.CurrentUser.OpenSubKey(path);
            if (key == null)
            {
                MicroEngActions.Log($"SelectionTreeRegistryProbe: {path} not found.");
                return;
            }

            var valueNames = key.GetValueNames() ?? Array.Empty<string>();
            var subKeyNames = key.GetSubKeyNames() ?? Array.Empty<string>();

            MicroEngActions.Log($"SelectionTreeRegistryProbe: {path} values={valueNames.Length}, subkeys={subKeyNames.Length}");

            foreach (var name in valueNames.OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
            {
                var valueKind = key.GetValueKind(name);
                var value = key.GetValue(name);
                var preview = FormatValuePreview(value);
                var displayName = string.IsNullOrEmpty(name) ? "(Default)" : name;
                MicroEngActions.Log($"SelectionTreeRegistryProbe: {path}::{displayName} [{valueKind}] {preview}");
            }

            foreach (var name in subKeyNames.OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
            {
                MicroEngActions.Log($"SelectionTreeRegistryProbe: {path} subkey={name}");
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
    }
}
