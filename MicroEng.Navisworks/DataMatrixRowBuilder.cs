using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace MicroEng.Navisworks
{
    internal class DataMatrixRowBuilder
    {
        public (List<DataMatrixAttributeDefinition> Attributes, List<DataMatrixRow> Rows) Build(ScrapeSession session)
        {
            var attributes = new Dictionary<string, DataMatrixAttributeDefinition>(StringComparer.OrdinalIgnoreCase);
            var rows = new Dictionary<string, DataMatrixRow>(StringComparer.OrdinalIgnoreCase);

            int order = 0;
            foreach (var prop in session.Properties)
            {
                var id = $"{prop.Category}|{prop.Name}";
                if (!attributes.ContainsKey(id))
                {
                    attributes[id] = new DataMatrixAttributeDefinition
                    {
                        Id = id,
                        Category = prop.Category,
                        PropertyName = prop.Name,
                        DisplayName = prop.Name,
                        DataType = MapType(prop.DataType),
                        IsEditable = false,
                        IsVisibleByDefault = true,
                        DefaultWidth = 160,
                        DisplayOrder = order++
                    };
                }
            }

            foreach (var entry in session.RawEntries)
            {
                var rowKey = !string.IsNullOrWhiteSpace(entry.ItemKey)
                    ? entry.ItemKey
                    : (!string.IsNullOrWhiteSpace(entry.ItemPath) ? entry.ItemPath : Guid.NewGuid().ToString());
                if (!rows.TryGetValue(rowKey, out var row))
                {
                    row = new DataMatrixRow
                    {
                        ItemKey = rowKey,
                        ElementDisplayName = entry.ItemPath
                    };
                    rows[rowKey] = row;
                }

                var attrId = $"{entry.Category}|{entry.Name}";
                if (!attributes.TryGetValue(attrId, out var attr))
                {
                    attr = new DataMatrixAttributeDefinition
                    {
                        Id = attrId,
                        Category = entry.Category,
                        PropertyName = entry.Name,
                        DisplayName = entry.Name,
                        DataType = MapType(entry.DataType),
                        IsEditable = false,
                        IsVisibleByDefault = true,
                        DefaultWidth = 160,
                        DisplayOrder = order++
                    };
                    attributes[attrId] = attr;
                }

                row.Values[attrId] = ConvertValue(entry.Value, attr.DataType);
            }

            return (attributes.Values.OrderBy(a => a.DisplayOrder).ToList(), rows.Values.ToList());
        }

        private Type MapType(string dataTypeName)
        {
            var name = dataTypeName ?? string.Empty;
            if (name.IndexOf("double", StringComparison.OrdinalIgnoreCase) >= 0) return typeof(double);
            if (name.IndexOf("int", StringComparison.OrdinalIgnoreCase) >= 0) return typeof(int);
            if (name.IndexOf("bool", StringComparison.OrdinalIgnoreCase) >= 0) return typeof(bool);
            if (name.IndexOf("date", StringComparison.OrdinalIgnoreCase) >= 0) return typeof(DateTime);
            return typeof(string);
        }

        private object ConvertValue(string value, Type target)
        {
            if (value == null) return null;
            try
            {
                if (target == typeof(int))
                {
                    if (int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var i)) return i;
                }
                if (target == typeof(double))
                {
                    if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
                }
                if (target == typeof(bool))
                {
                    if (bool.TryParse(value, out var b)) return b;
                }
                if (target == typeof(DateTime))
                {
                    if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt)) return dt;
                }
            }
            catch
            {
                // fall through
            }
            return value;
        }
    }
}
