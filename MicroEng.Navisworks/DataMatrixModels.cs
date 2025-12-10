using System;
using System.Collections.Generic;

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

    internal class DataMatrixSortDefinition
    {
        public string AttributeId { get; set; }
        public SortDirection Direction { get; set; }
        public int Priority { get; set; }
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

    internal class DataMatrixColumnFilter
    {
        public string AttributeId { get; set; }
        public DataMatrixFilterOperator Operator { get; set; }
        public string Value { get; set; }
        public List<string> ValuesList { get; set; } = new List<string>();
    }

    internal class DataMatrixViewPreset
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ScraperProfileName { get; set; }
        public List<string> VisibleAttributeIds { get; set; } = new List<string>();
        public List<DataMatrixSortDefinition> SortDefinitions { get; set; } = new List<DataMatrixSortDefinition>();
        public List<DataMatrixColumnFilter> Filters { get; set; } = new List<DataMatrixColumnFilter>();
    }
}
