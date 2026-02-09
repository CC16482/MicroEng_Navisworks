using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace MicroEng.Navisworks
{
    internal class DataMatrixRowBuilder
    {
        private static List<ScrapedProperty> GetDistinctProperties(ScrapeSession session)
        {
            var distinct = new List<ScrapedProperty>();
            var properties = session?.Properties;
            if (properties == null || properties.Count == 0)
            {
                return distinct;
            }

            var indexById = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var property in properties)
            {
                if (property == null)
                {
                    continue;
                }

                var id = BuildAttributeId(property.Category, property.Name);
                if (!indexById.TryGetValue(id, out var existingIndex))
                {
                    indexById[id] = distinct.Count;
                    distinct.Add(property);
                    continue;
                }

                if (IsPreferredProperty(property, distinct[existingIndex]))
                {
                    distinct[existingIndex] = property;
                }
            }

            return distinct;
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

            var whitelist = new List<string>();
            HashSet<string> whitelistSet = null;
            if (attributeWhitelist != null)
            {
                whitelistSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var id in attributeWhitelist)
                {
                    if (string.IsNullOrWhiteSpace(id))
                    {
                        continue;
                    }

                    if (whitelistSet.Add(id))
                    {
                        whitelist.Add(id);
                    }
                }

                if (whitelistSet.Count == 0)
                {
                    whitelistSet = null;
                }
            }

            var distinctProperties = GetDistinctProperties(session);
            var propById = new Dictionary<string, ScrapedProperty>(StringComparer.OrdinalIgnoreCase);
            foreach (var property in distinctProperties)
            {
                var id = BuildAttributeId(property.Category, property.Name);
                propById[id] = property;
            }

            var initialAttributeCapacity = whitelistSet != null
                ? whitelistSet.Count
                : Math.Max(distinctProperties.Count, 16);
            var attributes = new Dictionary<string, DataMatrixAttributeDefinition>(
                initialAttributeCapacity,
                StringComparer.OrdinalIgnoreCase);
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
                foreach (var p in distinctProperties)
                {
                    var id = BuildAttributeId(p.Category, p.Name);
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

            var estimatedRowCount = session.ItemsScanned > 0 ? session.ItemsScanned : 0;
            var rows = estimatedRowCount > 0
                ? new Dictionary<string, DataMatrixRow>(estimatedRowCount, StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, DataMatrixRow>(StringComparer.OrdinalIgnoreCase);
            var attrIdCache = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in session.RawEntries ?? Enumerable.Empty<RawEntry>())
            {
                if (entry == null)
                {
                    continue;
                }

                var rowKey = !string.IsNullOrWhiteSpace(entry.ItemKey)
                    ? entry.ItemKey
                    : (!string.IsNullOrWhiteSpace(entry.ItemPath) ? entry.ItemPath : Guid.NewGuid().ToString());

                if (itemKeyWhitelist != null && !itemKeyWhitelist.Contains(rowKey))
                {
                    continue;
                }

                var attrId = GetOrCreateAttributeId(attrIdCache, entry.Category, entry.Name);
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

        private static bool IsPreferredProperty(ScrapedProperty candidate, ScrapedProperty current)
        {
            if (candidate == null)
            {
                return false;
            }

            if (current == null)
            {
                return true;
            }

            if (candidate.ItemCount != current.ItemCount)
            {
                return candidate.ItemCount > current.ItemCount;
            }

            if (candidate.DistinctValueCount != current.DistinctValueCount)
            {
                return candidate.DistinctValueCount > current.DistinctValueCount;
            }

            return false;
        }

        private static string BuildAttributeId(string category, string name)
        {
            return string.Concat(category ?? string.Empty, "|", name ?? string.Empty);
        }

        private static string GetOrCreateAttributeId(
            IDictionary<string, Dictionary<string, string>> cache,
            string category,
            string name)
        {
            var normalizedCategory = category ?? string.Empty;
            var normalizedName = name ?? string.Empty;

            if (!cache.TryGetValue(normalizedCategory, out var byName))
            {
                byName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                cache[normalizedCategory] = byName;
            }

            if (!byName.TryGetValue(normalizedName, out var id))
            {
                id = BuildAttributeId(normalizedCategory, normalizedName);
                byName[normalizedName] = id;
            }

            return id;
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
