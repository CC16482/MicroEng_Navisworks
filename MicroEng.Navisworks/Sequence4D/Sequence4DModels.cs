using System;
using Autodesk.Navisworks.Api;

namespace MicroEng.Navisworks
{
    public enum Sequence4DOrdering
    {
        SelectionOrder = 0,
        DistanceToReference = 1,
        WorldXAscending = 2,
        WorldYAscending = 3,
        WorldZAscending = 4,
        PropertyValue = 5,
        Random = 6
    }

    public sealed class Sequence4DOptions
    {
        public ModelItemCollection SourceItems { get; set; } = new ModelItemCollection();

        public Sequence4DOrdering Ordering { get; set; } = Sequence4DOrdering.DistanceToReference;

        public ModelItem ReferenceItem { get; set; }

        public string PropertyPath { get; set; } = "Item|Name";

        public string SequenceName { get; set; } = "ME 4D Sequence";

        public string TaskNamePrefix { get; set; } = "Step ";

        public int ItemsPerTask { get; set; } = 1;

        public double DurationSeconds { get; set; } = 10.0;

        public double OverlapSeconds { get; set; } = 0.0;

        public DateTime StartDateTime { get; set; } = DateTime.Today.AddHours(8);

        public string SimulationTaskTypeName { get; set; } = "Construct";
    }
}
