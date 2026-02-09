using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using Autodesk.Navisworks.Api;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

namespace MicroEng.SelectionTreeCom
{
    [ComVisible(true)]
    [Guid("9D4B2B3D-36B5-4A4E-8B79-9D8E6B7D6C01")]
    [ProgId("MicroEng.TreeMapperSelectionTreePlugin")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class TreeMapperSelectionTreePlugin : ComApi.InwSelectionTreePlugin, ComApi.InwPlugin
    {
        private ComApi.InwPlugin_Site _site;

        public TreeMapperSelectionTreePlugin()
        {
            SelectionTreeComLog.Log("SelectionTreeCom: TreeMapperSelectionTreePlugin ctor");
        }

        public void AdviseSite(ComApi.InwPlugin_Site plugin_site)
        {
            _site = plugin_site;
            SelectionTreeComLog.Log("SelectionTreeCom: AdviseSite");
        }

        public bool iActivate()
        {
            SelectionTreeComLog.Log("SelectionTreeCom: iActivate");
            return true;
        }

        public bool iDeactivate()
        {
            SelectionTreeComLog.Log("SelectionTreeCom: iDeactivate");
            return true;
        }

        public int iGetNumParameters() => 0;

        public string iGetParameter(int ndx, ref object pData) => null;

        public bool iSetParameter(int ndx, object newVal) => false;

        public string ObjectName => "TreeMapperSelectionTreePlugin";

        public string iGetDisplayName() => "TreeMapper Selection Tree";

        public object iXtension(object vIn) => null;

        public void iAppInitialising()
        {
            SelectionTreeComLog.Log("SelectionTreeCom: iAppInitialising");
        }

        public void iAppTerminating()
        {
            SelectionTreeComLog.Log("SelectionTreeCom: iAppTerminating");
        }

        public void iDoCustomOption()
        {
            SelectionTreeComLog.Log("SelectionTreeCom: iDoCustomOption");
        }

        public void InitialisePlugin(ref int capbits, ref int ver)
        {
            SelectionTreeComLog.Log("SelectionTreeCom: InitialisePlugin");
            capbits = 0;
            ver = 1;
        }

        public string iGetUserString()
        {
            SelectionTreeComLog.Log("SelectionTreeCom: iGetUserString");
            return "TreeMapper";
        }

        public ComApi.InwOpSelectionTreeInterface iCreateInterface(ComApi.InwOpState State)
        {
            SelectionTreeComLog.Log("SelectionTreeCom: iCreateInterface");
            return new TreeMapperSelectionTreeInterface(State);
        }
    }

    [ComVisible(true)]
    [Guid("C1B33BE9-9A6D-4B44-9D06-7E8F7B0A2F02")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class TreeMapperSelectionTreeInterface : ComApi.InwOpSelectionTreeInterface
    {
        private const string MissingTreeMessage = "No published TreeMapper tree. Open TreeMapper and Publish.";
        private static readonly string PublishedTreePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MicroEng", "Navisworks", "TreeMapper", "PublishedTree.json");

        private readonly ComApi.InwOpState _state;
        private readonly object _sync = new object();
        private TreeMapperPublishedTree _tree;
        private Dictionary<string, TreeMapperPublishedNode> _nodeById;
        private Dictionary<string, List<string>> _leafCache;
        private DateTime _treeWriteUtc;
        private long _treeFileLength;
        private int? _pathIndexBase;
        private Dictionary<Guid, ModelItem> _itemIndex;
        private string _itemIndexDocument;
        private string _treeDocumentKey;

        public TreeMapperSelectionTreeInterface(ComApi.InwOpState state)
        {
            _state = state;
            SelectionTreeComLog.Log("SelectionTreeCom: TreeMapperSelectionTreeInterface ctor");
        }

        public int iGetNumRootChildren()
        {
            EnsureTreeLoaded();
            var count = _tree?.RootIds?.Count ?? 0;
            SelectionTreeComLog.Log($"SelectionTreeCom: iGetNumRootChildren={count}");
            return count;
        }

        public int iGetNumChildren(ComApi.InwUInt32Vector path, int user_handle)
        {
            var node = ResolveNode(path, user_handle, "iGetNumChildren");
            var count = node?.ChildIds?.Count ?? 0;
            SelectionTreeComLog.Log($"SelectionTreeCom: iGetNumChildren={count}");
            return count;
        }

        public string iGetName(ComApi.InwUInt32Vector path, int user_handle)
        {
            var node = ResolveNode(path, user_handle, "iGetName");
            var name = FormatNodeName(node);
            SelectionTreeComLog.Log($"SelectionTreeCom: iGetName={name}");
            return name;
        }

        public int iCreateHandle(ComApi.InwUInt32Vector path)
        {
            var indices = GetPathIndices(path, "iCreateHandle");
            if (indices.Length == 0) return 0;
            var handle = indices.Aggregate(17, (current, value) => unchecked(current * 31 + value));
            SelectionTreeComLog.Log($"SelectionTreeCom: iCreateHandle={handle}");
            return handle;
        }

        public void iDestroyHandle(ComApi.InwUInt32Vector path, int user_handle)
        {
        }

        public ComApi.nwESelTreeIcon iGetIcon(ComApi.InwUInt32Vector path, int user_handle)
        {
            var node = ResolveNode(path, user_handle, "iGetIcon");
            return MapIcon(node?.NodeType ?? TreeMapperNodeType.Collection);
        }

        public ComApi.nwESelTreeIcon iGetSelectedIcon(ComApi.InwUInt32Vector path, int user_handle)
        {
            var node = ResolveNode(path, user_handle, "iGetSelectedIcon");
            return MapIcon(node?.NodeType ?? TreeMapperNodeType.Collection);
        }

        public ComApi.nwESelTreeTextFormat iGetTextFormat(ComApi.InwUInt32Vector path, int user_handle) => ComApi.nwESelTreeTextFormat.eSelTreeTxtFmt_NORMAL;

        public void iGetSelection(ComApi.InwUInt32Vector path, int user_handle, ComApi.InwOpSelection selection)
        {
            try
            {
                var node = ResolveNode(path, user_handle, "iGetSelection");
                if (node == null) return;

                var leafKeys = GetLeafKeys(node);
                if (leafKeys.Count == 0) return;

                EnsureModelIndex();
                if (_itemIndex == null || _itemIndex.Count == 0) return;

                var paths = selection?.Paths();
                if (paths == null) return;

                paths.Clear();
                var added = 0;
                foreach (var key in leafKeys)
                {
                    if (!Guid.TryParse(key, out var guid)) continue;
                    if (!_itemIndex.TryGetValue(guid, out var item)) continue;
                    var oaPath = ComBridge.ToInwOaPath(item);
                    if (oaPath == null) continue;
                    paths.Add(oaPath);
                    added++;
                }

                SelectionTreeComLog.Log($"SelectionTreeCom: iGetSelection leafs={leafKeys.Count} added={added}");
            }
            catch (Exception ex)
            {
                SelectionTreeComLog.Log($"SelectionTreeCom: iGetSelection failed ({ex.Message})");
            }
        }

        public void iGetTextColor(ComApi.InwUInt32Vector path, int user_handle, ComApi.InwLVec3f txt_color)
        {
        }

        public void iOnBeginContextMenu(ComApi.InwUInt32Vector path, int user_handle)
        {
        }

        public void iOnEndContextMenu(ComApi.InwUInt32Vector path, int user_handle)
        {
        }

        public void iOnCollapsed(ComApi.InwUInt32Vector path, int user_handle)
        {
        }

        public void iOnLButtonDown(ComApi.InwUInt32Vector path, int user_handle, bool is_selected, bool is_shift, bool is_ctrl)
        {
        }

        private void EnsureTreeLoaded()
        {
            lock (_sync)
            {
                try
                {
                    if (!File.Exists(PublishedTreePath))
                    {
                        LoadMissingTree();
                        return;
                    }

                    var fileInfo = new FileInfo(PublishedTreePath);
                    var writeUtc = fileInfo.LastWriteTimeUtc;
                    var length = fileInfo.Length;
                    var currentDocKey = GetActiveDocumentKey();

                    if (_tree != null && writeUtc == _treeWriteUtc && length == _treeFileLength)
                    {
                        if (IsTreeForActiveDocument(_tree, currentDocKey))
                        {
                            _treeDocumentKey = currentDocKey;
                            return;
                        }

                        LoadMissingTree(BuildMismatchMessage(_tree));
                        return;
                    }

                    using var stream = File.OpenRead(PublishedTreePath);
                    var serializer = new DataContractJsonSerializer(typeof(TreeMapperPublishedTree));
                    var tree = serializer.ReadObject(stream) as TreeMapperPublishedTree;
                    if (tree == null)
                    {
                        LoadMissingTree();
                        return;
                    }

                    _tree = tree;
                    _treeWriteUtc = writeUtc;
                    _treeFileLength = length;
                    if (!IsTreeForActiveDocument(_tree, currentDocKey))
                    {
                        LoadMissingTree(BuildMismatchMessage(_tree));
                        return;
                    }

                    _treeDocumentKey = currentDocKey;
                    BuildNodeIndex();
                    _leafCache = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                    _pathIndexBase = null;
                    SelectionTreeComLog.Log($"SelectionTreeCom: loaded PublishedTree.json nodes={_tree.Nodes?.Count ?? 0}");
                }
                catch (Exception ex)
                {
                    SelectionTreeComLog.Log($"SelectionTreeCom: load PublishedTree.json failed ({ex.Message})");
                    LoadMissingTree();
                }
            }
        }

        private void LoadMissingTree()
        {
            LoadMissingTree(MissingTreeMessage);
        }

        private void LoadMissingTree(string message)
        {
            var node = new TreeMapperPublishedNode
            {
                Id = "missing",
                Name = message,
                NodeType = TreeMapperNodeType.Collection,
                Count = 0,
                ChildIds = new List<string>(),
                LeafItemKeys = new List<string>()
            };

            _tree = new TreeMapperPublishedTree
            {
                Version = 1,
                GeneratedUtc = DateTime.UtcNow,
                ProfileId = string.Empty,
                ProfileName = string.Empty,
                DocumentFile = string.Empty,
                DocumentFileKey = string.Empty,
                RootIds = new List<string> { node.Id },
                Nodes = new List<TreeMapperPublishedNode> { node }
            };

            BuildNodeIndex();
            _leafCache = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            _pathIndexBase = null;
            _treeDocumentKey = string.Empty;
        }

        private void BuildNodeIndex()
        {
            _nodeById = new Dictionary<string, TreeMapperPublishedNode>(StringComparer.OrdinalIgnoreCase);
            if (_tree?.Nodes == null) return;
            foreach (var node in _tree.Nodes)
            {
                if (string.IsNullOrWhiteSpace(node?.Id)) continue;
                if (_nodeById.ContainsKey(node.Id)) continue;
                _nodeById[node.Id] = node;
            }
        }

        private TreeMapperPublishedNode ResolveNode(ComApi.InwUInt32Vector path, int userHandle, string caller)
        {
            EnsureTreeLoaded();
            if (_tree?.RootIds == null || _tree.RootIds.Count == 0) return null;

            var indices = GetPathIndices(path, caller);
            if (indices.Length == 0) return null;

            if (_pathIndexBase.HasValue && TryResolveNode(indices, _pathIndexBase.Value, out var cached))
            {
                return cached;
            }

            if (TryResolveNode(indices, 0, out var zeroBased))
            {
                _pathIndexBase = 0;
                SelectionTreeComLog.Log("SelectionTreeCom: path index base=0");
                return zeroBased;
            }

            if (TryResolveNode(indices, 1, out var oneBased))
            {
                _pathIndexBase = 1;
                SelectionTreeComLog.Log("SelectionTreeCom: path index base=1");
                return oneBased;
            }

            SelectionTreeComLog.Log($"SelectionTreeCom: ResolveNode failed caller={caller} handle={userHandle} path=[{string.Join(",", indices)}]");
            return null;
        }

        private bool TryResolveNode(IReadOnlyList<int> indices, int baseIndex, out TreeMapperPublishedNode node)
        {
            node = null;
            if (_tree == null || _nodeById == null) return false;
            if (indices == null || indices.Count == 0) return false;

            var rootIndex = indices[0] - baseIndex;
            if (rootIndex < 0 || rootIndex >= _tree.RootIds.Count) return false;

            var currentId = _tree.RootIds[rootIndex];
            if (!_nodeById.TryGetValue(currentId, out var current)) return false;

            for (var depth = 1; depth < indices.Count; depth++)
            {
                if (current.ChildIds == null || current.ChildIds.Count == 0) return false;
                var childIndex = indices[depth] - baseIndex;
                if (childIndex < 0 || childIndex >= current.ChildIds.Count) return false;
                currentId = current.ChildIds[childIndex];
                if (!_nodeById.TryGetValue(currentId, out current)) return false;
            }

            node = current;
            return true;
        }

        private List<string> GetLeafKeys(TreeMapperPublishedNode node)
        {
            if (node == null) return new List<string>();
            if (!string.IsNullOrWhiteSpace(node.Id) && _leafCache != null && _leafCache.TryGetValue(node.Id, out var cached))
            {
                return cached;
            }

            var keys = new List<string>();
            if (node.LeafItemKeys != null && node.LeafItemKeys.Count > 0)
            {
                keys.AddRange(node.LeafItemKeys);
            }

            if (node.ChildIds != null && node.ChildIds.Count > 0)
            {
                foreach (var childId in node.ChildIds)
                {
                    if (string.IsNullOrWhiteSpace(childId)) continue;
                    if (!_nodeById.TryGetValue(childId, out var child)) continue;
                    keys.AddRange(GetLeafKeys(child));
                }
            }

            var distinct = keys
                .Where(k => !string.IsNullOrWhiteSpace(k))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (!string.IsNullOrWhiteSpace(node.Id))
            {
                _leafCache ??= new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                _leafCache[node.Id] = distinct;
            }

            return distinct;
        }

        private void EnsureModelIndex()
        {
            var doc = Application.ActiveDocument;
            if (doc == null) return;

            var docId = doc.FileName ?? string.Empty;
            if (_itemIndex != null && string.Equals(_itemIndexDocument, docId, StringComparison.OrdinalIgnoreCase)) return;

            var dict = new Dictionary<Guid, ModelItem>();
            foreach (var item in Traverse(doc.Models.RootItems))
            {
                try
                {
                    var guid = item.InstanceGuid;
                    if (guid == Guid.Empty) continue;
                    if (!dict.ContainsKey(guid)) dict.Add(guid, item);
                }
                catch
                {
                    // ignore item errors
                }
            }

            _itemIndex = dict;
            _itemIndexDocument = docId;
            SelectionTreeComLog.Log($"SelectionTreeCom: indexed model items={_itemIndex.Count}");
        }

        private static IEnumerable<ModelItem> Traverse(IEnumerable<ModelItem> items)
        {
            if (items == null) yield break;
            foreach (var item in items)
            {
                if (item == null) continue;
                yield return item;
                if (item.Children == null || !item.Children.Any()) continue;
                foreach (var child in Traverse(item.Children))
                {
                    yield return child;
                }
            }
        }

        private static string GetActiveDocumentKey()
        {
            try
            {
                var docFile = Application.ActiveDocument?.FileName ?? string.Empty;
                return BuildDocumentKey(docFile);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string BuildDocumentKey(string documentFile)
        {
            if (string.IsNullOrWhiteSpace(documentFile))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFileName(documentFile) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool IsTreeForActiveDocument(TreeMapperPublishedTree tree, string currentDocKey)
        {
            if (tree == null)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(tree.DocumentFileKey))
            {
                return false;
            }

            return string.Equals(tree.DocumentFileKey, currentDocKey ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildMismatchMessage(TreeMapperPublishedTree tree)
        {
            var key = tree?.DocumentFileKey ?? string.Empty;
            if (string.IsNullOrWhiteSpace(key))
            {
                return "TreeMapper tree is for a different model. Open TreeMapper and Publish.";
            }

            return $"TreeMapper tree was published for '{key}'. Publish for current model.";
        }

        private static int[] GetPathIndices(ComApi.InwUInt32Vector path, string caller)
        {
            if (path == null) return Array.Empty<int>();
            try
            {
                var data = path.ArrayData;
                if (data is Array array)
                {
                    if (array.Rank != 1)
                    {
                        SelectionTreeComLog.Log($"SelectionTreeCom: {caller} path ArrayData rank={array.Rank} type={array.GetType().FullName}");
                        return Array.Empty<int>();
                    }

                    var lower = array.GetLowerBound(0);
                    var upper = array.GetUpperBound(0);
                    var length = upper - lower + 1;
                    var indices = new int[length];
                    for (var i = 0; i < length; i++)
                    {
                        indices[i] = Convert.ToInt32(array.GetValue(lower + i));
                    }
                    SelectionTreeComLog.Log($"SelectionTreeCom: {caller} path len={indices.Length} lb={lower} data={string.Join(",", indices)}");
                    return indices;
                }
                if (data is System.Collections.IEnumerable enumerable)
                {
                    var list = new List<int>();
                    foreach (var value in enumerable)
                    {
                        list.Add(Convert.ToInt32(value));
                    }
                    if (list.Count > 0)
                    {
                        SelectionTreeComLog.Log($"SelectionTreeCom: {caller} path len={list.Count} data={string.Join(",", list)}");
                        return list.ToArray();
                    }
                }
                SelectionTreeComLog.Log($"SelectionTreeCom: {caller} path ArrayData type={(data == null ? "<null>" : data.GetType().FullName)}");
            }
            catch (Exception ex)
            {
                SelectionTreeComLog.Log($"SelectionTreeCom: {caller} path read failed ({ex.Message})");
            }

            return Array.Empty<int>();
        }

        private static ComApi.nwESelTreeIcon MapIcon(TreeMapperNodeType nodeType)
        {
            return nodeType switch
            {
                TreeMapperNodeType.Model => ComApi.nwESelTreeIcon.nwESelTreeIcon_MODEL_NORMAL,
                TreeMapperNodeType.Layer => ComApi.nwESelTreeIcon.nwESelTreeIcon_LAYER_NORMAL,
                TreeMapperNodeType.Group => ComApi.nwESelTreeIcon.nwESelTreeIcon_GROUP_NORMAL,
                TreeMapperNodeType.Composite => ComApi.nwESelTreeIcon.nwESelTreeIcon_COMPOSITE_NORMAL,
                TreeMapperNodeType.Insert => ComApi.nwESelTreeIcon.nwESelTreeIcon_INSERT_GROUP_NORMAL,
                TreeMapperNodeType.Instance => ComApi.nwESelTreeIcon.nwESelTreeIcon_INSERT_GEOMETRY_NORMAL,
                TreeMapperNodeType.Geometry => ComApi.nwESelTreeIcon.nwESelTreeIcon_GEOMETRY_NORMAL,
                TreeMapperNodeType.Collection => ComApi.nwESelTreeIcon.nwESelTreeIcon_COLLECTION_NORMAL,
                TreeMapperNodeType.Item => ComApi.nwESelTreeIcon.nwESelTreeIcon_GEOMETRY_NORMAL,
                _ => ComApi.nwESelTreeIcon.nwESelTreeIcon_UNKNOWN_NORMAL
            };
        }

        private static string FormatNodeName(TreeMapperPublishedNode node)
        {
            return StripTypePrefix(node?.Name ?? string.Empty);
        }

        private static string StripTypePrefix(string value)
        {
            var trimmed = value?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return string.Empty;
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
    }

    [DataContract]
    internal sealed class TreeMapperPublishedTree
    {
        [DataMember(Order = 1)] public int Version { get; set; }
        [DataMember(Order = 2)] public DateTime GeneratedUtc { get; set; }
        [DataMember(Order = 3)] public string ProfileId { get; set; }
        [DataMember(Order = 4)] public string ProfileName { get; set; }
        [DataMember(Order = 5)] public string DocumentFile { get; set; }
        [DataMember(Order = 6)] public string DocumentFileKey { get; set; }
        [DataMember(Order = 7)] public List<string> RootIds { get; set; } = new List<string>();
        [DataMember(Order = 8)] public List<TreeMapperPublishedNode> Nodes { get; set; } = new List<TreeMapperPublishedNode>();
    }

    [DataContract]
    internal sealed class TreeMapperPublishedNode
    {
        [DataMember(Order = 1)] public string Id { get; set; }
        [DataMember(Order = 2)] public string Name { get; set; }
        [DataMember(Order = 3)] public TreeMapperNodeType NodeType { get; set; }
        [DataMember(Order = 4)] public int Count { get; set; }
        [DataMember(Order = 5)] public List<string> ChildIds { get; set; } = new List<string>();
        [DataMember(Order = 6)] public List<string> LeafItemKeys { get; set; } = new List<string>();
    }

    internal enum TreeMapperNodeType
    {
        Model,
        Layer,
        Group,
        Composite,
        Geometry,
        Collection,
        Item,
        Insert,
        Instance
    }

    internal static class SelectionTreeComLog
    {
        private static readonly string LogPathPrimary = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MicroEng.Navisworks",
            "NavisErrors",
            "MicroEng.log");

        private static readonly string LogPathFallback = Path.Combine(Path.GetTempPath(), "MicroEng.SelectionTreeCom.log");

        public static void Log(string message)
        {
            var line = $"[{DateTime.Now:HH:mm:ss}] {message}";
            WriteLine(LogPathPrimary, line);
            WriteLine(LogPathFallback, line);
        }

        private static void WriteLine(string path, string line)
        {
            try
            {
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                File.AppendAllLines(path, new[] { line });
            }
            catch
            {
                // Never throw from COM entry points; logging is best-effort.
            }
        }
    }
}
