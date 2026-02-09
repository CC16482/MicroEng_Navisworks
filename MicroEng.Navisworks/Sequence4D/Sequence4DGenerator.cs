using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Timeliner;

namespace MicroEng.Navisworks
{
    internal static class Sequence4DGenerator
    {
        public static int GenerateTimelinerSequence(Sequence4DOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            var doc = Autodesk.Navisworks.Api.Application.ActiveDocument
                      ?? throw new InvalidOperationException("No active Navisworks document.");

            if (options.SourceItems == null || options.SourceItems.Count == 0)
            {
                throw new InvalidOperationException("No source items. Capture a selection first.");
            }

            var timeliner = TimelinerDocumentExtensions.GetTimeliner(doc);

            var ordered = OrderItems(options);

            var itemsPerTask = Math.Max(1, options.ItemsPerTask);
            var duration = Math.Max(0.1, options.DurationSeconds);
            var overlap = Math.Max(0.0, options.OverlapSeconds);
            if (overlap >= duration)
            {
                overlap = Math.Max(0.0, duration - 0.01);
            }

            var step = duration - overlap;

            var taskCount = (int)Math.Ceiling(ordered.Count / (double)itemsPerTask);
            if (taskCount <= 0)
            {
                return 0;
            }

            var parent = new TimelinerTask
            {
                DisplayName = string.IsNullOrWhiteSpace(options.SequenceName) ? "ME 4D Sequence" : options.SequenceName
            };
            TrySetTaskTypeName(parent, options.SimulationTaskTypeName);

            var seqStart = options.StartDateTime;

            for (int i = 0; i < taskCount; i++)
            {
                var chunk = ordered.Skip(i * itemsPerTask).Take(itemsPerTask).ToList();
                var selectionItems = new ModelItemCollection();
                foreach (var item in chunk)
                {
                    selectionItems.Add(item);
                }

                var start = seqStart.AddSeconds(i * step);
                var end = start.AddSeconds(duration);

                var child = new TimelinerTask
                {
                    DisplayName = BuildStepName(options.TaskNamePrefix, i + 1)
                };

                TrySetTaskTypeName(child, options.SimulationTaskTypeName);
                TrySetPlannedDates(child, start, end);

                var selection = BuildSelection(selectionItems);
                child.Selection.CopyFrom(selection);

                parent.Children.Add(child);
            }

            TrySetPlannedDates(
                parent,
                seqStart,
                seqStart.AddSeconds((taskCount - 1) * step + duration));

            TryRecalculateSummaryDates(parent);

            using (var tx = doc.BeginTransaction("MicroEng 4D Sequence: Generate"))
            {
                if (!TryTaskAddCopy(timeliner, parent))
                {
                    AddByTasksCopyFrom(timeliner, parent);
                }

                tx.Commit();
            }

            return taskCount;
        }

        public static int DeleteTimelinerSequence(string sequenceName)
        {
            if (string.IsNullOrWhiteSpace(sequenceName))
            {
                throw new ArgumentException("Sequence name is required.", nameof(sequenceName));
            }

            var doc = Autodesk.Navisworks.Api.Application.ActiveDocument
                      ?? throw new InvalidOperationException("No active Navisworks document.");

            var timeliner = TimelinerDocumentExtensions.GetTimeliner(doc);
            var removed = 0;

            using (var tx = doc.BeginTransaction("MicroEng 4D Sequence: Delete"))
            {
                var root = timeliner.TasksRoot;
                for (int i = root.Children.Count - 1; i >= 0; i--)
                {
                    if (root.Children[i] is TimelinerTask task)
                    {
                        if (string.Equals(task.DisplayName, sequenceName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (!TryTaskRemoveAt(timeliner, root, i))
                            {
                                RemoveByTasksCopyFrom(timeliner, sequenceName);
                            }
                            removed++;
                        }
                    }
                }

                tx.Commit();
            }

            return removed;
        }

        private static List<ModelItem> OrderItems(Sequence4DOptions options)
        {
            var items = options.SourceItems.Cast<ModelItem>().Where(x => x != null).ToList();

            switch (options.Ordering)
            {
                case Sequence4DOrdering.SelectionOrder:
                    return items;

                case Sequence4DOrdering.Random:
                {
                    var rng = new Random();
                    return items.OrderBy(_ => rng.Next()).ToList();
                }

                case Sequence4DOrdering.DistanceToReference:
                {
                    if (options.ReferenceItem == null)
                    {
                        throw new InvalidOperationException("Distance ordering requires a reference item. Use 'Set Reference'.");
                    }

                    var refPt = Sequence4DModelItemUtils.GetBoundingBoxCenter(options.ReferenceItem);
                    return items
                        .Select(mi => new
                        {
                            Item = mi,
                            Dist = Sequence4DModelItemUtils.Distance(Sequence4DModelItemUtils.GetBoundingBoxCenter(mi), refPt)
                        })
                        .OrderBy(x => x.Dist)
                        .Select(x => x.Item)
                        .ToList();
                }

                case Sequence4DOrdering.WorldXAscending:
                    return items.OrderBy(mi => Sequence4DModelItemUtils.GetBoundingBoxCenter(mi).X).ToList();

                case Sequence4DOrdering.WorldYAscending:
                    return items.OrderBy(mi => Sequence4DModelItemUtils.GetBoundingBoxCenter(mi).Y).ToList();

                case Sequence4DOrdering.WorldZAscending:
                    return items.OrderBy(mi => Sequence4DModelItemUtils.GetBoundingBoxCenter(mi).Z).ToList();

                case Sequence4DOrdering.PropertyValue:
                {
                    var path = options.PropertyPath ?? string.Empty;
                    return items
                        .Select(mi => new { Item = mi, Key = Sequence4DModelItemUtils.GetPropertyValueString(mi, path) ?? string.Empty })
                        .OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
                        .Select(x => x.Item)
                        .ToList();
                }

                default:
                    return items;
            }
        }

        private static string BuildStepName(string prefix, int index1)
        {
            var p = string.IsNullOrWhiteSpace(prefix) ? "Step " : prefix;
            return $"{p}{index1:000}";
        }

        private static void TrySetPlannedDates(TimelinerTask task, DateTime start, DateTime end)
        {
            TrySetProperty(task, new[] { "PlannedStartDate", "PlannedStart", "PlannedStartDateTime" }, start);
            TrySetProperty(task, new[] { "PlannedEndDate", "PlannedEnd", "PlannedEndDateTime" }, end);
        }

        private static void TrySetTaskTypeName(TimelinerTask task, string typeName)
        {
            if (string.IsNullOrWhiteSpace(typeName))
            {
                return;
            }

            TrySetProperty(task, new[] { "SimulationTaskTypeName", "TaskTypeName", "SimulationTaskType" }, typeName);
        }

        private static void TrySetProperty(object obj, IEnumerable<string> candidateNames, object value)
        {
            var type = obj.GetType();
            foreach (var name in candidateNames)
            {
                var prop = type.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                if (prop == null || !prop.CanWrite)
                {
                    continue;
                }

                try
                {
                    object converted = value;
                    if (value != null && !prop.PropertyType.IsInstanceOfType(value))
                    {
                        converted = Convert.ChangeType(value, prop.PropertyType);
                    }

                    prop.SetValue(obj, converted, null);
                    return;
                }
                catch
                {
                    // try next
                }
            }
        }

        private static void TryRecalculateSummaryDates(TimelinerTask parent)
        {
            try
            {
                var method = typeof(DocumentTimeliner).GetMethod(
                    "TaskSummaryHierarchyRecalculateDates",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new[] { typeof(TimelinerTask) },
                    null);

                method?.Invoke(null, new object[] { parent });
            }
            catch
            {
                // Non-fatal; depends on task writability in some versions.
            }
        }

        private static Selection BuildSelection(ModelItemCollection items)
        {
            var ctor = typeof(Selection).GetConstructor(new[] { typeof(ModelItemCollection) });
            if (ctor != null)
            {
                return (Selection)ctor.Invoke(new object[] { items });
            }

            var selection = new Selection();
            var copyFrom = typeof(Selection).GetMethod("CopyFrom", new[] { typeof(ModelItemCollection) });
            if (copyFrom != null)
            {
                copyFrom.Invoke(selection, new object[] { items });
                return selection;
            }

            throw new InvalidOperationException("Unable to construct a Selection from ModelItemCollection on this Navisworks version.");
        }

        private static bool TryTaskAddCopy(DocumentTimeliner timeliner, TimelinerTask rootTask)
        {
            try
            {
                var method = timeliner.GetType().GetMethod("TaskAddCopy", new[] { typeof(TimelinerTask) });
                if (method == null)
                {
                    return false;
                }

                method.Invoke(timeliner, new object[] { rootTask });
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryTaskRemoveAt(DocumentTimeliner timeliner, GroupItem parent, int index)
        {
            try
            {
                var method = timeliner.GetType().GetMethod("TaskRemoveAt", new[] { typeof(GroupItem), typeof(int) });
                if (method == null)
                {
                    return false;
                }

                method.Invoke(timeliner, new object[] { parent, index });
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void AddByTasksCopyFrom(DocumentTimeliner timeliner, TimelinerTask parentTask)
        {
            var rootCopy = timeliner.TasksRoot.CreateCopy() as GroupItem;
            if (rootCopy == null)
            {
                throw new InvalidOperationException("Could not copy Timeliner TasksRoot.");
            }

            rootCopy.Children.Add(parentTask);
            timeliner.TasksCopyFrom(rootCopy.Children);
        }

        private static void RemoveByTasksCopyFrom(DocumentTimeliner timeliner, string sequenceName)
        {
            var rootCopy = timeliner.TasksRoot.CreateCopy() as GroupItem;
            if (rootCopy == null)
            {
                throw new InvalidOperationException("Could not copy Timeliner TasksRoot.");
            }

            for (int i = rootCopy.Children.Count - 1; i >= 0; i--)
            {
                if (rootCopy.Children[i] is TimelinerTask task &&
                    string.Equals(task.DisplayName, sequenceName, StringComparison.OrdinalIgnoreCase))
                {
                    rootCopy.Children.RemoveAt(i);
                }
            }

            timeliner.TasksCopyFrom(rootCopy.Children);
        }
    }
}
