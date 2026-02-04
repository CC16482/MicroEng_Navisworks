using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace MicroEng.Navisworks.TreeMapper
{
    internal static class TreeMapperEngine
    {
        public static List<TreeMapperPreviewNode> BuildPreview(
            ScrapeSession session,
            IReadOnlyList<TreeMapperLevel> levels,
            CancellationToken token)
        {
            var roots = BuildNodes(session, levels, token, out var activeLevels);
            if (roots == null || activeLevels.Count == 0)
            {
                return new List<TreeMapperPreviewNode>();
            }

            return ConvertNodes(roots, activeLevels, 0);
        }

        public static TreeMapperPublishedTree BuildPublishedTree(
            ScrapeSession session,
            IReadOnlyList<TreeMapperLevel> levels,
            string profileId,
            string profileName,
            CancellationToken token)
        {
            var published = new TreeMapperPublishedTree
            {
                ProfileId = profileId,
                ProfileName = profileName,
                DocumentFile = session?.DocumentFile ?? string.Empty,
                DocumentFileKey = session?.DocumentFileKey ?? string.Empty,
                GeneratedUtc = DateTime.UtcNow
            };

            var roots = BuildNodes(session, levels, token, out var activeLevels);
            if (roots == null || activeLevels.Count == 0)
            {
                return published;
            }

            var nodes = new List<TreeMapperPublishedNode>();
            var rootIds = new List<string>();

            foreach (var node in OrderNodes(roots, activeLevels, 0))
            {
                token.ThrowIfCancellationRequested();
                rootIds.Add(FlattenNode(node, activeLevels, 0, nodes, token));
            }

            published.RootIds = rootIds;
            published.Nodes = nodes;
            return published;
        }

        private static Dictionary<string, BuilderNode> BuildNodes(
            ScrapeSession session,
            IReadOnlyList<TreeMapperLevel> levels,
            CancellationToken token,
            out List<TreeMapperLevel> activeLevels)
        {
            activeLevels = new List<TreeMapperLevel>();
            if (session?.RawEntries == null || levels == null || levels.Count == 0)
            {
                return null;
            }

            activeLevels = levels
                .Where(l => !string.IsNullOrWhiteSpace(l.Category) && !string.IsNullOrWhiteSpace(l.PropertyName))
                .ToList();

            if (activeLevels.Count == 0)
            {
                return null;
            }

            var keyToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < activeLevels.Count; i++)
            {
                var level = activeLevels[i];
                var key = BuildKey(level.Category, level.PropertyName);
                if (!keyToIndex.ContainsKey(key))
                {
                    keyToIndex[key] = i;
                }
            }

            var itemValues = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            var allItemKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in session.RawEntries)
            {
                token.ThrowIfCancellationRequested();

                if (entry == null)
                {
                    continue;
                }

                var itemKey = entry.ItemKey ?? string.Empty;
                if (string.IsNullOrWhiteSpace(itemKey))
                {
                    continue;
                }

                allItemKeys.Add(itemKey);

                var key = BuildKey(entry.Category, entry.Name);
                if (!keyToIndex.TryGetValue(key, out var levelIndex))
                {
                    continue;
                }

                if (!itemValues.TryGetValue(itemKey, out var values))
                {
                    values = new string[activeLevels.Count];
                    itemValues[itemKey] = values;
                }

                if (string.IsNullOrWhiteSpace(values[levelIndex]))
                {
                    values[levelIndex] = Normalize(entry.Value, entry.DataType);
                }
            }

            var roots = new Dictionary<string, BuilderNode>(StringComparer.OrdinalIgnoreCase);

            foreach (var itemKey in allItemKeys)
            {
                token.ThrowIfCancellationRequested();

                if (!itemValues.TryGetValue(itemKey, out var values))
                {
                    values = new string[activeLevels.Count];
                    itemValues[itemKey] = values;
                }

                for (var i = 0; i < activeLevels.Count; i++)
                {
                    if (string.IsNullOrWhiteSpace(values[i]))
                    {
                        values[i] = activeLevels[i].MissingLabel ?? "(Missing)";
                    }
                }

                var current = roots;
                for (var i = 0; i < activeLevels.Count; i++)
                {
                    var level = activeLevels[i];
                    var value = values[i] ?? "";

                    if (!current.TryGetValue(value, out var node))
                    {
                        node = new BuilderNode
                        {
                            Name = value,
                            NodeType = level.NodeType
                        };
                        current[value] = node;
                    }

                    node.Count++;

                    if (i == activeLevels.Count - 1)
                    {
                        node.LeafItemKeys.Add(itemKey);
                    }

                    current = node.Children;
                }
            }

            return roots;
        }

        private static List<TreeMapperPreviewNode> ConvertNodes(
            Dictionary<string, BuilderNode> nodes,
            IReadOnlyList<TreeMapperLevel> levels,
            int depth)
        {
            var results = OrderNodes(nodes, levels, depth)
                .Select(n => new TreeMapperPreviewNode
                {
                    Name = n.Name,
                    Count = n.Count,
                    NodeType = n.NodeType,
                    Children = ConvertNodes(n.Children, levels, depth + 1)
                })
                .ToList();

            return results;
        }

        private static IEnumerable<BuilderNode> OrderNodes(
            Dictionary<string, BuilderNode> nodes,
            IReadOnlyList<TreeMapperLevel> levels,
            int depth)
        {
            if (nodes == null)
            {
                return Enumerable.Empty<BuilderNode>();
            }

            return nodes.Values
                .OrderBy(n => n, new BuilderNodeComparer(GetSortMode(levels, depth)))
                .ToList();
        }

        private static string FlattenNode(
            BuilderNode node,
            IReadOnlyList<TreeMapperLevel> levels,
            int depth,
            List<TreeMapperPublishedNode> sink,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            var id = Guid.NewGuid().ToString("N");
            var childIds = new List<string>();

            foreach (var child in OrderNodes(node.Children, levels, depth + 1))
            {
                token.ThrowIfCancellationRequested();
                childIds.Add(FlattenNode(child, levels, depth + 1, sink, token));
            }

            var leafKeys = childIds.Count == 0
                ? node.LeafItemKeys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase).ToList()
                : new List<string>();

            var published = new TreeMapperPublishedNode
            {
                Id = id,
                Name = node.Name,
                NodeType = node.NodeType,
                Count = node.Count,
                ChildIds = childIds,
                LeafItemKeys = leafKeys
            };

            sink.Add(published);
            return id;
        }

        private static TreeMapperSortMode GetSortMode(IReadOnlyList<TreeMapperLevel> levels, int depth)
        {
            if (levels == null || levels.Count == 0)
            {
                return TreeMapperSortMode.Alpha;
            }

            var index = Math.Max(0, Math.Min(depth, levels.Count - 1));
            return levels[index].SortMode;
        }

        private static string BuildKey(string category, string property)
        {
            return $"{category ?? string.Empty}\u001F{property ?? string.Empty}";
        }

        private static string Normalize(string value, string dataType)
        {
            var trimmed = value?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return string.Empty;
            }

            if (!string.IsNullOrWhiteSpace(dataType))
            {
                var dt = dataType.Trim();
                if (!string.IsNullOrWhiteSpace(dt))
                {
                    if (trimmed.StartsWith(dt + ":", StringComparison.OrdinalIgnoreCase))
                    {
                        return trimmed.Substring(dt.Length + 1).Trim();
                    }

                    var lastDot = dt.LastIndexOf('.');
                    var dtShort = lastDot >= 0 && lastDot < dt.Length - 1 ? dt.Substring(lastDot + 1) : dt;
                    if (!string.Equals(dtShort, dt, StringComparison.OrdinalIgnoreCase)
                        && trimmed.StartsWith(dtShort + ":", StringComparison.OrdinalIgnoreCase))
                    {
                        return trimmed.Substring(dtShort.Length + 1).Trim();
                    }

                    var systemPrefix = "System." + dtShort + ":";
                    if (trimmed.StartsWith(systemPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        return trimmed.Substring(systemPrefix.Length).Trim();
                    }
                }
            }

            var colon = trimmed.IndexOf(':');
            if (colon > 0 && colon < 32)
            {
                var prefix = trimmed.Substring(0, colon);
                if (KnownTypePrefixes.Contains(prefix))
                {
                    return trimmed.Substring(colon + 1).Trim();
                }
            }

            return trimmed;
        }

        private static readonly HashSet<string> KnownTypePrefixes = new(StringComparer.OrdinalIgnoreCase)
        {
            "Boolean",
            "Bool",
            "Byte",
            "SByte",
            "Int16",
            "Int32",
            "Int64",
            "UInt16",
            "UInt32",
            "UInt64",
            "Single",
            "Double",
            "Decimal",
            "DateTime",
            "String",
            "Guid",
            "System.Boolean",
            "System.Byte",
            "System.SByte",
            "System.Int16",
            "System.Int32",
            "System.Int64",
            "System.UInt16",
            "System.UInt32",
            "System.UInt64",
            "System.Single",
            "System.Double",
            "System.Decimal",
            "System.DateTime",
            "System.String",
            "System.Guid"
        };

        private sealed class BuilderNode
        {
            public string Name;
            public int Count;
            public TreeMapperNodeType NodeType;
            public Dictionary<string, BuilderNode> Children = new(StringComparer.OrdinalIgnoreCase);
            public HashSet<string> LeafItemKeys = new(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class BuilderNodeComparer : IComparer<BuilderNode>
        {
            private readonly TreeMapperSortMode _mode;

            public BuilderNodeComparer(TreeMapperSortMode mode)
            {
                _mode = mode;
            }

            public int Compare(BuilderNode x, BuilderNode y)
            {
                if (ReferenceEquals(x, y))
                {
                    return 0;
                }

                if (x == null) return -1;
                if (y == null) return 1;

                var a = x.Name ?? string.Empty;
                var b = y.Name ?? string.Empty;

                if (_mode == TreeMapperSortMode.NumericLike)
                {
                    if (double.TryParse(a, NumberStyles.Any, CultureInfo.InvariantCulture, out var da)
                        && double.TryParse(b, NumberStyles.Any, CultureInfo.InvariantCulture, out var db))
                    {
                        return da.CompareTo(db);
                    }
                }

                return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
