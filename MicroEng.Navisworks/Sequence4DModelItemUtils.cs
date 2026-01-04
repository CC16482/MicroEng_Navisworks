using System;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Data;

namespace MicroEng.Navisworks
{
    internal static class Sequence4DModelItemUtils
    {
        public static Point3D GetBoundingBoxCenter(ModelItem item)
        {
            if (item == null)
            {
                return new Point3D(0, 0, 0);
            }

            var bounds = item.BoundingBox();
            if (bounds == null || bounds.IsEmpty)
            {
                return new Point3D(0, 0, 0);
            }

            var min = bounds.Min;
            var max = bounds.Max;
            return new Point3D(
                (min.X + max.X) * 0.5,
                (min.Y + max.Y) * 0.5,
                (min.Z + max.Z) * 0.5);
        }

        public static double Distance(Point3D a, Point3D b)
        {
            var dx = a.X - b.X;
            var dy = a.Y - b.Y;
            var dz = a.Z - b.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        public static string GetPropertyValueString(ModelItem item, string propertyPath)
        {
            if (item == null || string.IsNullOrWhiteSpace(propertyPath))
            {
                return string.Empty;
            }

            var parts = propertyPath.Split('|');
            if (parts.Length != 2)
            {
                return string.Empty;
            }

            var categoryName = parts[0].Trim();
            var propertyName = parts[1].Trim();

            try
            {
                foreach (PropertyCategory category in item.PropertyCategories)
                {
                    if (!string.Equals(category.DisplayName, categoryName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    foreach (DataProperty prop in category.Properties)
                    {
                        if (!string.Equals(prop.DisplayName, propertyName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        return prop.Value == null ? string.Empty : prop.Value.ToString();
                    }
                }
            }
            catch
            {
                // Some items can have unusual property stacks; treat as empty.
            }

            return string.Empty;
        }
    }
}
