using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace MicroEng.Navisworks.TreeMapper
{
    [DataContract]
    internal sealed class TreeMapperProfileStoreData
    {
        [DataMember(Order = 1)] public int Version { get; set; } = 1;
        [DataMember(Order = 2)] public string ActiveProfileId { get; set; }
        [DataMember(Order = 3)] public List<TreeMapperProfile> Profiles { get; set; } = new();
    }

    [DataContract]
    internal sealed class TreeMapperProfile
    {
        [DataMember(Order = 1)] public string Id { get; set; } = Guid.NewGuid().ToString();
        [DataMember(Order = 2)] public string Name { get; set; } = "TreeMapper";
        [DataMember(Order = 3)] public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;
        [DataMember(Order = 4)] public List<TreeMapperLevel> Levels { get; set; } = new();
    }

    [DataContract]
    internal sealed class TreeMapperLevel
    {
        [DataMember(Order = 1)] public TreeMapperNodeType NodeType { get; set; } = TreeMapperNodeType.Group;
        [DataMember(Order = 2)] public string Category { get; set; }
        [DataMember(Order = 3)] public string PropertyName { get; set; }
        [DataMember(Order = 4)] public string MissingLabel { get; set; } = "(Missing)";
        [DataMember(Order = 5)] public TreeMapperSortMode SortMode { get; set; } = TreeMapperSortMode.Alpha;
    }

    internal enum TreeMapperNodeType
    {
        Model,
        Layer,
        Group,
        Composite,
        Geometry,
        Collection,
        Item
    }

    internal enum TreeMapperSortMode
    {
        Alpha,
        NumericLike
    }

    internal sealed class TreeMapperPreviewNode
    {
        public string Name { get; set; }
        public int Count { get; set; }
        public TreeMapperNodeType NodeType { get; set; }
        public List<TreeMapperPreviewNode> Children { get; set; } = new();

        public string DisplayLabel => $"{Name} ({Count})";
    }

    [DataContract]
    internal sealed class TreeMapperPublishedTree
    {
        [DataMember(Order = 1)] public int Version { get; set; } = 1;
        [DataMember(Order = 2)] public DateTime GeneratedUtc { get; set; } = DateTime.UtcNow;
        [DataMember(Order = 3)] public string ProfileId { get; set; }
        [DataMember(Order = 4)] public string ProfileName { get; set; }
        [DataMember(Order = 5)] public string DocumentFile { get; set; }
        [DataMember(Order = 6)] public string DocumentFileKey { get; set; }
        [DataMember(Order = 7)] public List<string> RootIds { get; set; } = new();
        [DataMember(Order = 8)] public List<TreeMapperPublishedNode> Nodes { get; set; } = new();
    }

    [DataContract]
    internal sealed class TreeMapperPublishedNode
    {
        [DataMember(Order = 1)] public string Id { get; set; }
        [DataMember(Order = 2)] public string Name { get; set; }
        [DataMember(Order = 3)] public TreeMapperNodeType NodeType { get; set; }
        [DataMember(Order = 4)] public int Count { get; set; }
        [DataMember(Order = 5)] public List<string> ChildIds { get; set; } = new();
        [DataMember(Order = 6)] public List<string> LeafItemKeys { get; set; } = new();
    }
}
