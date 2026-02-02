using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace MicroEng.Navisworks
{
    internal class DataMatrixAttributeDefinition
    {
        public string Id { get; set; } // e.g. $"{Category}|{PropertyName}"
        public string Category { get; set; }
        public string PropertyName { get; set; }
        public string DisplayName { get; set; }
        public Type DataType { get; set; } = typeof(string);
        public bool IsEditable { get; set; }
        public bool IsVisibleByDefault { get; set; } = true;
        public int DefaultWidth { get; set; } = 140;
        public int DisplayOrder { get; set; }
    }

    internal class DataMatrixRow
    {
        public string ItemKey { get; set; }
        public string ElementDisplayName { get; set; }
        public Autodesk.Navisworks.Api.ModelItem ModelItem { get; set; }
        public Dictionary<string, object> Values { get; set; } = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
    }

    internal enum SortDirection
    {
        None,
        Ascending,
        Descending
    }

    [DataContract]
    internal class DataMatrixSortDefinition
    {
        [DataMember(Order = 1)] public string AttributeId { get; set; }
        [DataMember(Order = 2)] public SortDirection Direction { get; set; }
        [DataMember(Order = 3)] public int Priority { get; set; }
    }

    internal enum DataMatrixFilterOperator
    {
        None,
        Contains,
        Equals,
        NotEquals,
        GreaterThan,
        LessThan,
        Blank,
        NotBlank,
        InList
    }

    [DataContract]
    internal class DataMatrixColumnFilter
    {
        [DataMember(Order = 1)] public string AttributeId { get; set; }
        [DataMember(Order = 2)] public DataMatrixFilterOperator Operator { get; set; }
        [DataMember(Order = 3)] public string Value { get; set; }
        [DataMember(Order = 4)] public List<string> ValuesList { get; set; } = new List<string>();
    }

    [DataContract]
    internal enum DataMatrixScopeKind
    {
        [EnumMember] EntireSession,
        [EnumMember] CurrentSelection,
        [EnumMember] SingleItem,
        [EnumMember] SelectionSet,
        [EnumMember] SearchSet
    }

    [DataContract]
    internal enum DataMatrixJsonlMode
    {
        [EnumMember] ItemDocuments,
        [EnumMember] RawRows
    }

    [DataContract]
    internal class DataMatrixViewPreset
    {
        [DataMember(Order = 1)] public string Id { get; set; }
        [DataMember(Order = 2)] public string Name { get; set; }
        [DataMember(Order = 3)] public string ScraperProfileName { get; set; }

        [DataMember(Order = 4)] public DataMatrixScopeKind ScopeKind { get; set; } = DataMatrixScopeKind.EntireSession;
        [DataMember(Order = 5)] public string ScopeSetName { get; set; }
        [DataMember(Order = 6)] public bool JoinMultiValues { get; set; } = false;
        [DataMember(Order = 7)] public string MultiValueSeparator { get; set; } = "; ";

        [DataMember(Order = 8)] public List<string> VisibleAttributeIds { get; set; } = new List<string>();
        [DataMember(Order = 9)] public List<DataMatrixSortDefinition> SortDefinitions { get; set; } = new List<DataMatrixSortDefinition>();
        [DataMember(Order = 10)] public List<DataMatrixColumnFilter> Filters { get; set; } = new List<DataMatrixColumnFilter>();

        [DataMember(Order = 11)] public double ElementColumnWidth { get; set; } = 180;
        [DataMember(Order = 12)] public Dictionary<string, double> ColumnWidths { get; set; } = new Dictionary<string, double>();
    }
}
