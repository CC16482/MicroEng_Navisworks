using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;

namespace MicroEng.Navisworks
{
    [Plugin("MicroEng.Dev.SelectionTreeProbe", "MENG",
        DisplayName = "MicroEng Dev: Selection Tree Extensibility Probe",
        ToolTip = "Dumps runtime type information to discover whether Navisworks exposes a custom Selection Tree dropdown/tab API.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class DevSelectionTreeProbeAddIn : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();
            MicroEngActions.Log("SelectionTreeProbe: start");

            try
            {
                var diagDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MicroEng.Navisworks",
                    "Diagnostics");

                Directory.CreateDirectory(diagDir);

                var filePath = Path.Combine(
                    diagDir,
                    $"SelectionTreeProbe_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

                var sb = new StringBuilder();
                sb.AppendLine("=== MicroEng Selection Tree Extensibility Probe ===");
                sb.AppendLine($"Timestamp: {DateTime.Now:u}");
                sb.AppendLine($"Process: {GetProcessPath()}");
                sb.AppendLine();

                var assemblies = AppDomain.CurrentDomain.GetAssemblies()
                    .OrderBy(a => a.FullName, StringComparer.OrdinalIgnoreCase)
                    .ToArray();

                sb.AppendLine("=== Loaded assemblies ===");
                foreach (var a in assemblies)
                {
                    string loc = "<dynamic>";
                    try
                    {
                        loc = string.IsNullOrWhiteSpace(a.Location) ? "<dynamic>" : a.Location;
                    }
                    catch
                    {
                        // ignore
                    }

                    sb.AppendLine($"- {a.FullName}");
                    sb.AppendLine($"  Location: {loc}");
                }

                sb.AppendLine();
                sb.AppendLine("=== Candidate types (name match) ===");

                string[] typeNeedles =
                {
                    "SelectionTree",
                    "ViewTree",
                    "ModelTree",
                    "TreeView",
                    "TreeTab",
                    "TreeProvider"
                };

                foreach (var a in assemblies)
                {
                    Type[] types;
                    try
                    {
                        types = a.GetTypes();
                    }
                    catch (ReflectionTypeLoadException rtle)
                    {
                        types = rtle.Types.Where(t => t != null).ToArray();
                    }
                    catch
                    {
                        continue;
                    }

                    foreach (var t in types)
                    {
                        var fullName = t.FullName ?? t.Name;
                        if (!typeNeedles.Any(n => fullName.IndexOf(n, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            continue;
                        }

                        sb.AppendLine();
                        sb.AppendLine($"TYPE: {fullName}");
                        sb.AppendLine($"  Assembly: {a.GetName().Name}");
                        sb.AppendLine($"  Base: {t.BaseType?.FullName ?? "<none>"}");

                        try
                        {
                            var ifaces = t.GetInterfaces()
                                .Select(i => i.FullName)
                                .OrderBy(s => s)
                                .ToArray();
                            sb.AppendLine($"  Interfaces: {(ifaces.Length == 0 ? "<none>" : string.Join(", ", ifaces))}");
                        }
                        catch
                        {
                            sb.AppendLine("  Interfaces: <error>");
                        }

                        try
                        {
                            var props = t.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static)
                                .Select(p => $"{p.PropertyType.Name} {p.Name}")
                                .OrderBy(s => s)
                                .ToArray();
                            sb.AppendLine("  Public Properties:");
                            if (props.Length == 0) sb.AppendLine("    <none>");
                            foreach (var p in props) sb.AppendLine($"    {p}");
                        }
                        catch
                        {
                            sb.AppendLine("  Public Properties: <error>");
                        }

                        try
                        {
                            var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly)
                                .Where(m => !m.IsSpecialName)
                                .Select(MethodSig)
                                .OrderBy(s => s)
                                .ToArray();
                            sb.AppendLine("  Public Methods (declared only):");
                            if (methods.Length == 0) sb.AppendLine("    <none>");
                            foreach (var m in methods) sb.AppendLine($"    {m}");
                        }
                        catch
                        {
                            sb.AppendLine("  Public Methods: <error>");
                        }
                    }
                }

                sb.AppendLine();
                sb.AppendLine("=== Candidate methods (global scan) ===");
                sb.AppendLine("Looking for method names containing 'SelectionTree' and ('Add' or 'Register')");

                foreach (var a in assemblies)
                {
                    Type[] types;
                    try
                    {
                        types = a.GetTypes();
                    }
                    catch (ReflectionTypeLoadException rtle)
                    {
                        types = rtle.Types.Where(t => t != null).ToArray();
                    }
                    catch
                    {
                        continue;
                    }

                    foreach (var t in types)
                    {
                        MethodInfo[] methods;
                        try
                        {
                            methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);
                        }
                        catch
                        {
                            continue;
                        }

                        foreach (var m in methods)
                        {
                            var name = m.Name ?? "";
                            if (name.IndexOf("SelectionTree", StringComparison.OrdinalIgnoreCase) < 0)
                            {
                                continue;
                            }

                            var hasAdd = name.IndexOf("Add", StringComparison.OrdinalIgnoreCase) >= 0;
                            var hasReg = name.IndexOf("Register", StringComparison.OrdinalIgnoreCase) >= 0;
                            if (!hasAdd && !hasReg)
                            {
                                continue;
                            }

                            sb.AppendLine($"{t.FullName} :: {MethodSig(m)}");
                        }
                    }
                }

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                MicroEngActions.Log($"SelectionTreeProbe wrote: {filePath}");

                MessageBox.Show(
                    $"Selection Tree probe complete.\n\nOutput:\n{filePath}",
                    "MicroEng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return 0;
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"SelectionTreeProbe failed: {ex}");
                MessageBox.Show(
                    $"Selection Tree probe failed:\n{ex.Message}",
                    "MicroEng",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return 0;
            }
        }

        private static string MethodSig(MethodInfo m)
        {
            try
            {
                var pars = m.GetParameters()
                    .Select(p => $"{p.ParameterType.Name} {p.Name}")
                    .ToArray();
                return $"{m.ReturnType.Name} {m.Name}({string.Join(", ", pars)})";
            }
            catch
            {
                return m.Name ?? "<unknown>";
            }
        }

        private static string GetProcessPath()
        {
            try
            {
                return Process.GetCurrentProcess().MainModule?.FileName ?? "<unknown>";
            }
            catch
            {
                return "<unknown>";
            }
        }
    }
}
