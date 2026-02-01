using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace MicroEng.Navisworks
{
    internal class DataMatrixRowBuilder
    {
        private static IEnumerable<ScrapedProperty> GetDistinctProperties(ScrapeSession session)
        {
            return (session?.Properties ?? Enumerable.Empty<ScrapedProperty>())
                .GroupBy(p => $"{p.Category}|{p.Name}", StringComparer.OrdinalIgnoreCase)
                .Select(g => g
                    .OrderByDescending(p => p.ItemCount)
                    .ThenByDescending(p => p.DistinctValueCount)
                    .First());
        }

        public List<DataMatrixAttributeDefinition> BuildAttributeCatalog(ScrapeSession session)
        {
            var list = new List<DataMatrixAttributeDefinition>();
            if (session?.Properties == null) return list;

            var order = 0;
            foreach (var p in GetDistinctProperties(session))
            {
                list.Add(new DataMatrixAttributeDefinition
                {
                    Id = $"{p.Category}|{p.Name}",
                    Category = p.Category,
                    PropertyName = p.Name,
                    DisplayName = p.Name,
                    DataType = MapType(p.DataType),
                    IsEditable = false,
                    IsVisibleByDefault = true,
                    DefaultWidth = 160,
                    DisplayOrder = order++
                });
            }

            return list;
        }

        public (List<DataMatrixAttributeDefinition> Attributes, List<DataMatrixRow> Rows) Build(ScrapeSession session)
        {
            return Build(session, null, null, joinMultiValues: false, multiValueSeparator: "; ");
        }

        public (List<DataMatrixAttributeDefinition> Attributes, List<DataMatrixRow> Rows) Build(
            ScrapeSession session,
            IList<string> attributeWhitelist,
            ISet<string> itemKeyWhitelist,
            bool joinMultiValues,
            string multiValueSeparator)
        {
            if (session == null)
            {
                return (new List<DataMatrixAttributeDefinition>(), new List<DataMatrixRow>());
            }

            var whitelist = (attributeWhitelist ?? new List<string>())
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var whitelistSet = whitelist.Count > 0
                ? new HashSet<string>(whitelist, StringComparer.OrdinalIgnoreCase)
                : null;

            var propById = GetDistinctProperties(session)
                .ToDictionary(p => $"{p.Category}|{p.Name}", StringComparer.OrdinalIgnoreCase);

            var attributes = new Dictionary<string, DataMatrixAttributeDefinition>(StringComparer.OrdinalIgnoreCase);
            var order = 0;

            if (whitelistSet != null)
            {
                foreach (var id in whitelist)
                {
                    if (propById.TryGetValue(id, out var p))
                    {
                        attributes[id] = new DataMatrixAttributeDefinition
                        {
                            Id = id,
                            Category = p.Category,
                            PropertyName = p.Name,
                            DisplayName = p.Name,
                            DataType = MapType(p.DataType),
                            IsEditable = false,
                            IsVisibleByDefault = true,
                            DefaultWidth = 160,
                            DisplayOrder = order++
                        };
                    }
                }
            }
            else
            {
                foreach (var p in session.Properties ?? Enumerable.Empty<ScrapedProperty>())
                {
                    var id = $"{p.Category}|{p.Name}";
                    attributes[id] = new DataMatrixAttributeDefinition
                    {
                        Id = id,
                        Category = p.Category,
                        PropertyName = p.Name,
                        DisplayName = p.Name,
                        DataType = MapType(p.DataType),
                        IsEditable = false,
                        IsVisibleByDefault = true,
                        DefaultWidth = 160,
                        DisplayOrder = order++
                    };
                }
            }

            var rows = new Dictionary<string, DataMatrixRow>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in session.RawEntries ?? Enumerable.Empty<RawEntry>())
            {
                var rowKey = !string.IsNullOrWhiteSpace(entry.ItemKey)
                    ? entry.ItemKey
                    : (!string.IsNullOrWhiteSpace(entry.ItemPath) ? entry.ItemPath : Guid.NewGuid().ToString());

                if (itemKeyWhitelist != null && !itemKeyWhitelist.Contains(rowKey))
                {
                    continue;
                }

                var attrId = $"{entry.Category}|{entry.Name}";
                if (whitelistSet != null && !whitelistSet.Contains(attrId))
                {
                    continue;
                }

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

                if (!rows.TryGetValue(rowKey, out var row))
                {
                    row = new DataMatrixRow
                    {
                        ItemKey = rowKey,
                        ElementDisplayName = entry.ItemPath
                    };
                    rows[rowKey] = row;
                }

                var converted = ConvertValue(entry.Value, attr.DataType);

                if (joinMultiValues
                    && row.Values.TryGetValue(attrId, out var existing)
                    && existing != null && converted != null
                    && !Equals(existing, converted))
                {
                    row.Values[attrId] = (existing.ToString() ?? string.Empty)
                                         + (multiValueSeparator ?? "; ")
                                         + (converted.ToString() ?? string.Empty);
                }
                else
                {
                    row.Values[attrId] = converted;
                }
            }

            List<DataMatrixAttributeDefinition> orderedAttrs;
            if (whitelistSet != null)
            {
                orderedAttrs = whitelist.Where(id => attributes.ContainsKey(id)).Select(id => attributes[id]).ToList();
            }
            else
            {
                orderedAttrs = attributes.Values.OrderBy(a => a.DisplayOrder).ToList();
            }

            return (orderedAttrs, rows.Values.ToList());
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
