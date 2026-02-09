using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.DocumentParts;

namespace MicroEng.Navisworks.ViewpointsGenerator
{
    internal static class ViewpointsGeneratorNavisworksService
    {
        private sealed class SelectionSetAccessor
        {
            public PropertyInfo SearchProperty;
            public PropertyInfo ExplicitItemsProperty;
        }

        private sealed class BoundingBoxAccessor
        {
            public PropertyInfo BoundingBoxProperty;
            public MethodInfo BoundingBoxMethod;
        }

        private static readonly ConcurrentDictionary<Type, SelectionSetAccessor> SelectionSetAccessorCache =
            new ConcurrentDictionary<Type, SelectionSetAccessor>();

        private static readonly ConcurrentDictionary<Type, BoundingBoxAccessor> BoundingBoxAccessorCache =
            new ConcurrentDictionary<Type, BoundingBoxAccessor>();

        public static List<SelectionSetPickerItem> LoadSelectionSets(Document doc)
        {
            var results = new List<SelectionSetPickerItem>();
            if (doc?.SelectionSets?.RootItem == null)
            {
                return results;
            }

            void Walk(GroupItem parent, string prefix)
            {
                foreach (var child in parent.Children)
                {
                    if (child is FolderItem folder)
                    {
                        Walk(folder, CombinePath(prefix, folder.DisplayName));
                    }
                    else if (child is SelectionSet set)
                    {
                        results.Add(new SelectionSetPickerItem
                        {
                            Enabled = true,
                            Set = set,
                            Path = CombinePath(prefix, set.DisplayName)
                        });
                    }
                }
            }

            Walk(doc.SelectionSets.RootItem, "");
            return results.OrderBy(x => x.Path).ToList();
        }

        public static List<ViewpointPlanItem> BuildPlan(
            Document doc,
            ViewpointsGeneratorSettings settings,
            IEnumerable<SelectionSetPickerItem> selectedSets)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            switch (settings.SourceMode)
            {
                case ViewpointsSourceMode.CurrentSelection:
                    return BuildFromCurrentSelection(doc, settings);
                case ViewpointsSourceMode.SelectionSets:
                    return BuildFromSelectionSets(doc, settings, selectedSets);
                case ViewpointsSourceMode.PropertyGroups:
                default:
                    return new List<ViewpointPlanItem>();
            }
        }

        private static List<ViewpointPlanItem> BuildFromCurrentSelection(Document doc, ViewpointsGeneratorSettings settings)
        {
            var items = doc.CurrentSelection?.SelectedItems;
            int count = items?.Count ?? 0;

            return new List<ViewpointPlanItem>
            {
                new ViewpointPlanItem
                {
                    Enabled = true,
                    Name = MakeName(settings, "Selection"),
                    Source = "Current Selection",
                    ItemCount = count,
                    ResolveItems = () => doc.CurrentSelection?.SelectedItems ?? new ModelItemCollection()
                }
            };
        }

        private static List<ViewpointPlanItem> BuildFromSelectionSets(
            Document doc,
            ViewpointsGeneratorSettings settings,
            IEnumerable<SelectionSetPickerItem> selectedSets)
        {
            var plan = new List<ViewpointPlanItem>();
            var sets = (selectedSets ?? Enumerable.Empty<SelectionSetPickerItem>())
                .Where(s => s != null && s.Enabled && s.Set != null)
                .ToList();

            foreach (var s in sets)
            {
                var resolved = ResolveSelectionSetItems(doc, s.Set);
                int count = resolved?.Count ?? 0;

                string shortName = string.IsNullOrWhiteSpace(s.Path) ? s.Set.DisplayName : s.Path;
                plan.Add(new ViewpointPlanItem
                {
                    Enabled = true,
                    Name = MakeName(settings, shortName),
                    Source = $"SelectionSet: {shortName}",
                    ItemCount = count,
                    ResolveItems = () => ResolveSelectionSetItems(doc, s.Set) ?? new ModelItemCollection()
                });
            }

            return plan;
        }

        public static void GenerateSavedViewpoints(
            Document doc,
            ViewpointsGeneratorSettings settings,
            IReadOnlyList<ViewpointPlanItem> planItems,
            Action<int, int, string> progress)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            if (planItems == null) throw new ArgumentNullException(nameof(planItems));

            var enabled = planItems.Where(p => p != null && p.Enabled).ToList();
            int total = enabled.Count;

            var folder = EnsureSavedViewpointFolder(doc, settings.OutputFolderPath);

            for (int i = 0; i < total; i++)
            {
                var plan = enabled[i];
                progress?.Invoke(i, total, plan.Name);

                var items = plan.ResolveItems?.Invoke() ?? new ModelItemCollection();
                if (items.Count == 0)
                {
                    continue;
                }

                if (!TryComputeBounds(items, out var bbox))
                {
                    continue;
                }

                var vp = CreateFittedViewpoint(doc, bbox, settings);
                var saved = CreateSavedViewpoint(vp, plan.Name);

                doc.SavedViewpoints.AddCopy(folder, saved);
            }

            progress?.Invoke(total, total, "Done");
        }

        private static GroupItem EnsureSavedViewpointFolder(Document doc, string folderPath)
        {
            GroupItem current = doc.SavedViewpoints.RootItem;

            var parts = (folderPath ?? "")
                .Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            foreach (var part in parts)
            {
                var existing = FindChildFolder(current, part);
                if (existing != null)
                {
                    current = existing;
                    continue;
                }

                var f = new FolderItem { DisplayName = part };
                doc.SavedViewpoints.AddCopy(current, f);

                existing = FindChildFolder(current, part);
                current = existing ?? current;
            }

            return current;
        }

        private static FolderItem FindChildFolder(GroupItem parent, string name)
        {
            foreach (var child in parent.Children)
            {
                if (child is FolderItem folder && string.Equals(folder.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                {
                    return folder;
                }
            }

            return null;
        }

        private static string CombinePath(string a, string b)
        {
            if (string.IsNullOrWhiteSpace(a)) return b ?? "";
            if (string.IsNullOrWhiteSpace(b)) return a ?? "";
            return a.TrimEnd('/', '\\') + "/" + b.TrimStart('/', '\\');
        }

        private static string MakeName(ViewpointsGeneratorSettings settings, string baseName)
        {
            var name = (baseName ?? "").Trim();
            if (name.Length == 0) name = "View";

            var p = (settings.NamePrefix ?? "").Trim();
            var s = (settings.NameSuffix ?? "").Trim();

            string outName = name;
            if (p.Length > 0) outName = p + outName;
            if (s.Length > 0) outName = outName + s;
            return outName;
        }

        private static ModelItemCollection ResolveSelectionSetItems(Document doc, SelectionSet set)
        {
            var accessor = GetSelectionSetAccessor(set?.GetType());
            var search = accessor?.SearchProperty?.GetValue(set) as Search;
            if (search != null)
            {
                return search.FindAll(doc, false);
            }

            var explicitItems = accessor?.ExplicitItemsProperty?.GetValue(set) as ModelItemCollection;
            if (explicitItems != null)
            {
                return explicitItems;
            }

            return new ModelItemCollection();
        }

        private static SavedViewpoint CreateSavedViewpoint(Viewpoint viewpoint, string name)
        {
            try
            {
                var sv = new SavedViewpoint(viewpoint) { DisplayName = name };
                return sv;
            }
            catch
            {
                var sv = new SavedViewpoint { DisplayName = name };
                var vpProp = sv.GetType().GetProperty("Viewpoint", BindingFlags.Instance | BindingFlags.Public);
                vpProp?.SetValue(sv, viewpoint);
                return sv;
            }
        }

        private static Viewpoint CreateFittedViewpoint(Document doc, BoundingBox3D bbox, ViewpointsGeneratorSettings settings)
        {
            Viewpoint vp = doc.CurrentViewpoint;
            var copyMethod = vp.GetType().GetMethod("CreateCopy", BindingFlags.Instance | BindingFlags.Public);
            if (copyMethod != null)
            {
                vp = (Viewpoint)copyMethod.Invoke(vp, null);
            }

            FitCameraToBox(vp, bbox, settings.Direction, settings.Projection, settings.FitMarginFactor);
            return vp;
        }

        private static void FitCameraToBox(Viewpoint vp, BoundingBox3D bbox, ViewDirectionPreset dir, ProjectionMode proj, double marginFactor)
        {
            var min = bbox.Min;
            var max = bbox.Max;

            var cx = (min.X + max.X) * 0.5;
            var cy = (min.Y + max.Y) * 0.5;
            var cz = (min.Z + max.Z) * 0.5;

            var sx = (max.X - min.X);
            var sy = (max.Y - min.Y);
            var sz = (max.Z - min.Z);

            double diag = Math.Sqrt(sx * sx + sy * sy + sz * sz);
            if (diag < 1e-6) diag = 1.0;

            double distance = diag * (1.2 + marginFactor);

            var (vx, vy, vz, ux, uy, uz) = GetPresetVectors(dir);

            double vlen = Math.Sqrt(vx * vx + vy * vy + vz * vz);
            if (vlen < 1e-9) { vx = 1; vy = 1; vz = 1; vlen = Math.Sqrt(3); }
            vx /= vlen; vy /= vlen; vz /= vlen;

            var pos = new Point3D(cx + vx * distance, cy + vy * distance, cz + vz * distance);
            var tgt = new Point3D(cx, cy, cz);
            var up = new Vector3D(ux, uy, uz);

            var camProp = vp.GetType().GetProperty("Camera", BindingFlags.Instance | BindingFlags.Public);
            var camObj = camProp?.GetValue(vp);
            if (camObj == null)
            {
                return;
            }

            SetIfExists(camObj, "Position", pos);
            SetIfExists(camObj, "Target", tgt);
            SetIfExists(camObj, "UpVector", up);

            var projProp = camObj.GetType().GetProperty("Projection", BindingFlags.Instance | BindingFlags.Public);
            if (projProp != null && projProp.PropertyType.IsEnum)
            {
                string want = (proj == ProjectionMode.Orthographic) ? "Orthographic" : "Perspective";
                try
                {
                    object enumVal = Enum.GetValues(projProp.PropertyType)
                        .Cast<object>()
                        .FirstOrDefault(v => string.Equals(v.ToString(), want, StringComparison.OrdinalIgnoreCase));

                    if (enumVal != null)
                    {
                        projProp.SetValue(camObj, enumVal);
                    }
                }
                catch
                {
                }
            }

            try { camProp?.SetValue(vp, camObj); } catch { }
        }

        private static (double vx, double vy, double vz, double ux, double uy, double uz) GetPresetVectors(ViewDirectionPreset dir)
        {
            switch (dir)
            {
                case ViewDirectionPreset.Top: return (0, 0, 1, 0, 1, 0);
                case ViewDirectionPreset.Bottom: return (0, 0, -1, 0, 1, 0);
                case ViewDirectionPreset.Front: return (0, -1, 0, 0, 0, 1);
                case ViewDirectionPreset.Back: return (0, 1, 0, 0, 0, 1);
                case ViewDirectionPreset.Left: return (-1, 0, 0, 0, 0, 1);
                case ViewDirectionPreset.Right: return (1, 0, 0, 0, 0, 1);
                case ViewDirectionPreset.IsoSW: return (-1, -1, 1, 0, 0, 1);
                case ViewDirectionPreset.IsoSE:
                default:
                    return (1, -1, 1, 0, 0, 1);
            }
        }

        private static void SetIfExists(object obj, string propName, object value)
        {
            var p = obj.GetType().GetProperty(propName, BindingFlags.Instance | BindingFlags.Public);
            if (p != null && p.CanWrite)
            {
                try { p.SetValue(obj, value); } catch { }
            }
        }

        private static bool TryComputeBounds(ModelItemCollection items, out BoundingBox3D bbox)
        {
            if (TryGetBoundingBox(items, out bbox))
            {
                return true;
            }

            bool any = false;
            BoundingBox3D acc = default;

            foreach (var it in items)
            {
                if (it == null) continue;
                if (!TryGetBoundingBox(it, out var b))
                {
                    continue;
                }

                if (!any)
                {
                    acc = b;
                    any = true;
                }
                else
                {
                    acc = Union(acc, b);
                }
            }

            bbox = acc;
            return any;
        }

        private static bool TryGetBoundingBox(object obj, out BoundingBox3D bbox)
        {
            bbox = default;

            var accessor = GetBoundingBoxAccessor(obj?.GetType());

            var prop = accessor?.BoundingBoxProperty;
            if (prop != null)
            {
                try
                {
                    bbox = (BoundingBox3D)prop.GetValue(obj);
                    return true;
                }
                catch
                {
                }
            }

            var method = accessor?.BoundingBoxMethod;
            if (method != null)
            {
                try
                {
                    bbox = (BoundingBox3D)method.Invoke(obj, null);
                    return true;
                }
                catch
                {
                }
            }

            return false;
        }

        private static SelectionSetAccessor GetSelectionSetAccessor(Type type)
        {
            if (type == null)
            {
                return null;
            }

            return SelectionSetAccessorCache.GetOrAdd(type, t => new SelectionSetAccessor
            {
                SearchProperty = t.GetProperty("Search", BindingFlags.Instance | BindingFlags.Public),
                ExplicitItemsProperty = t.GetProperty("ExplicitModelItems", BindingFlags.Instance | BindingFlags.Public)
            });
        }

        private static BoundingBoxAccessor GetBoundingBoxAccessor(Type type)
        {
            if (type == null)
            {
                return null;
            }

            return BoundingBoxAccessorCache.GetOrAdd(type, t =>
            {
                var accessor = new BoundingBoxAccessor
                {
                    BoundingBoxProperty = t.GetProperty("BoundingBox", BindingFlags.Instance | BindingFlags.Public),
                    BoundingBoxMethod = t.GetMethod("BoundingBox", BindingFlags.Instance | BindingFlags.Public, null, Type.EmptyTypes, null)
                };

                if (accessor.BoundingBoxProperty?.PropertyType != typeof(BoundingBox3D))
                {
                    accessor.BoundingBoxProperty = null;
                }

                if (accessor.BoundingBoxMethod?.ReturnType != typeof(BoundingBox3D))
                {
                    accessor.BoundingBoxMethod = null;
                }

                return accessor;
            });
        }

        private static BoundingBox3D Union(BoundingBox3D a, BoundingBox3D b)
        {
            var min = new Point3D(
                Math.Min(a.Min.X, b.Min.X),
                Math.Min(a.Min.Y, b.Min.Y),
                Math.Min(a.Min.Z, b.Min.Z));

            var max = new Point3D(
                Math.Max(a.Max.X, b.Max.X),
                Math.Max(a.Max.Y, b.Max.Y),
                Math.Max(a.Max.Z, b.Max.Z));

            return new BoundingBox3D(min, max);
        }
    }
}
