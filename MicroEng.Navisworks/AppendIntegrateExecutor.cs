using System;
using System.Linq;
using Autodesk.Navisworks.Api;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace MicroEng.Navisworks
{
    internal class AppendIntegrateResult
    {
        public int ItemsProcessed { get; set; }
        public int PropertiesCreated { get; set; }
        public int PropertiesUpdated { get; set; }
        public int PropertiesDeleted { get; set; }
        public string Message { get; set; }
    }

    internal class AppendIntegrateExecutor
    {
        private readonly AppendIntegrateTemplate _template;
        private readonly Action<string> _log;

        public AppendIntegrateExecutor(AppendIntegrateTemplate template, Action<string> log)
        {
            _template = template;
            _log = log;
        }

        public AppendIntegrateResult Execute()
        {
            var result = new AppendIntegrateResult();
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                result.Message = "No active document.";
                return result;
            }

            var targets = ResolveTargets(doc).ToList();
            if (!targets.Any())
            {
                result.Message = _template.ApplyToSelectionOnly
                    ? "No items selected."
                    : "No items found in the model.";
                return result;
            }

            foreach (var item in targets)
            {
                ProcessItem(item, result);
                result.ItemsProcessed++;
            }

            result.Message = $"Append & Integrate Data: {result.ItemsProcessed} items processed, " +
                             $"{result.PropertiesCreated} properties created, " +
                             $"{result.PropertiesUpdated} updated, " +
                             $"{result.PropertiesDeleted} removed.";
            return result;
        }

        private void ProcessItem(ModelItem item, AppendIntegrateResult result)
        {
            foreach (var row in _template.Rows.Where(r => r.Enabled))
            {
                var value = ComputeValue(row, item);
                var applied = ApplyProperty(item, row, value, out var created, out var updated, out var deleted);
                if (!applied) continue;

                if (created) result.PropertiesCreated++;
                if (updated) result.PropertiesUpdated++;
                if (deleted) result.PropertiesDeleted++;
            }
        }

        private string ComputeValue(AppendIntegrateRow row, ModelItem item)
        {
            switch (row.Mode)
            {
                case AppendValueMode.StaticValue:
                    return row.StaticOrExpressionValue ?? string.Empty;
                case AppendValueMode.FromProperty:
                    return ReadProperty(item, row.SourcePropertyPath);
                case AppendValueMode.Expression:
                    // Expression parsing placeholder; treat as literal for now.
                    return row.StaticOrExpressionValue ?? string.Empty;
                default:
                    return string.Empty;
            }
        }

        private string ReadProperty(ModelItem item, string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;

            var tokens = path.Split('|');
            var categoryKey = tokens.Length > 0 ? tokens[0] : string.Empty;
            var propKey = tokens.Length > 1 ? tokens[1] : string.Empty;

            foreach (var category in item.PropertyCategories)
            {
                if (category == null) continue;
                if (!KeyMatch(category.Name, categoryKey) && !KeyMatch(category.DisplayName, categoryKey))
                {
                    continue;
                }

                foreach (var prop in category.Properties)
                {
                    if (KeyMatch(prop.Name, propKey) || KeyMatch(prop.DisplayName, propKey))
                    {
                        return prop.Value?.ToDisplayString() ?? string.Empty;
                    }
                }
            }

            return string.Empty;
        }

        private bool ApplyProperty(ModelItem item, AppendIntegrateRow row, string value,
            out bool created, out bool updated, out bool deleted)
        {
            created = updated = deleted = false;
            if (string.IsNullOrWhiteSpace(row.TargetPropertyName))
            {
                return false;
            }

            var state = ComBridge.State;
            var path = ComBridge.ToInwOaPath(item);
            var propertyNode = (ComApi.InwGUIPropertyNode2)state.GetGUIPropertyNode(path, _template.CreateTargetTabIfMissing);
            if (propertyNode == null)
            {
                return false;
            }

            ComApi.InwOaPropertyVec propertyVec = null;
            var tabExists = false;
            try
            {
                var getMethod = propertyNode.GetType().GetMethod("GetUserDefined");
                propertyVec = getMethod?.Invoke(propertyNode, new object[] { 0, _template.TargetTabName }) as ComApi.InwOaPropertyVec;
                tabExists = propertyVec != null;
            }
            catch
            {
                // ignore read failures
            }

            if (propertyVec == null)
            {
                if (!_template.CreateTargetTabIfMissing)
                {
                    return false;
                }

                propertyVec = (ComApi.InwOaPropertyVec)state.ObjectFactory(
                    ComApi.nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
            }
            else if (!_template.UpdateExistingTargetTab)
            {
                return false;
            }

            var existingProp = FindProperty(propertyVec, row.TargetPropertyName);

            if (string.IsNullOrWhiteSpace(value) && _template.DeletePropertyIfBlank)
            {
                if (existingProp != null)
                {
                    RemoveProperty(propertyVec, existingProp);
                    deleted = true;
                }
            }
            else
            {
                if (existingProp == null)
                {
                    existingProp = (ComApi.InwOaProperty)state.ObjectFactory(
                        ComApi.nwEObjectType.eObjectType_nwOaProperty, null, null);
                    existingProp.name = row.TargetPropertyName;
                    existingProp.UserName = row.TargetPropertyName;
                    propertyVec.Properties().Add(existingProp);
                    created = true;
                }
                else
                {
                    updated = true;
                }

                existingProp.value = ApplyOption(value, row);
            }

            try
            {
                propertyNode.SetUserDefined(0, _template.TargetTabName, _template.TargetTabName, propertyVec);
            }
            catch (Exception ex)
            {
                _log?.Invoke($"Failed to set property '{row.TargetPropertyName}' on '{item.DisplayName}': {ex.Message}");
                return false;
            }

            if (_template.DeleteTargetTabIfAllBlank && !HasAnyProperties(propertyVec) && tabExists)
            {
                try
                {
                    var removeMethod = propertyNode.GetType().GetMethod("RemoveUserDefined");
                    removeMethod?.Invoke(propertyNode, new object[] { 0, _template.TargetTabName });
                }
                catch
                {
                    // If RemoveUserDefined is unavailable, leaving an empty tab is acceptable for V1.
                }
            }

            return true;
        }

        private static ComApi.InwOaProperty FindProperty(ComApi.InwOaPropertyVec vec, string propertyName)
        {
            foreach (ComApi.InwOaProperty prop in vec.Properties())
            {
                if (prop == null) continue;
                if (KeyMatch(prop.name, propertyName) || KeyMatch(prop.UserName, propertyName))
                {
                    return prop;
                }
            }

            return null;
        }

        private static void RemoveProperty(ComApi.InwOaPropertyVec vec, ComApi.InwOaProperty prop)
        {
            var props = vec.Properties();
            var index = 1; // COM collections are 1-based
            foreach (ComApi.InwOaProperty p in props)
            {
                if (p == prop)
                {
                    props.Remove(index);
                    return;
                }
                index++;
            }
        }

        private static bool HasAnyProperties(ComApi.InwOaPropertyVec vec)
        {
            foreach (ComApi.InwOaProperty _ in vec.Properties())
            {
                return true;
            }

            return false;
        }

        private static bool KeyMatch(string value, string key)
        {
            return string.Equals(value ?? string.Empty, key ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        private static object ApplyOption(string value, AppendIntegrateRow row)
        {
            var trimmed = value ?? string.Empty;
            switch (row.Option)
            {
                case AppendValueOption.FormatAsText:
                case AppendValueOption.ConvertToDecimal:
                case AppendValueOption.ConvertToInteger:
                case AppendValueOption.ConvertToDate:
                case AppendValueOption.SumGroupProperty:
                case AppendValueOption.UseParentProperty:
                case AppendValueOption.UseParentRevitProperty:
                case AppendValueOption.ReadAllPropertiesFromTab:
                case AppendValueOption.ParseExcelFormula:
                case AppendValueOption.PerformCalculation:
                case AppendValueOption.ConvertFromRevit:
                    return trimmed;
                default:
                    return trimmed;
            }
        }

        private System.Collections.Generic.IEnumerable<ModelItem> ResolveTargets(Document doc)
        {
            System.Collections.Generic.IEnumerable<ModelItem> items;
            if (_template.ApplyToSelectionOnly)
            {
                items = doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>() ??
                        System.Linq.Enumerable.Empty<ModelItem>();
            }
            else
            {
                items = Traverse(doc.Models.RootItems);
            }

            return items.Where(item =>
            {
                var isGroup = item.Children != null && item.Children.Any();
                return _template.ApplyPropertyTo switch
                {
                    ApplyPropertyTarget.Items => !isGroup,
                    ApplyPropertyTarget.Groups => isGroup,
                    _ => true
                };
            });
        }

        private System.Collections.Generic.IEnumerable<ModelItem> Traverse(System.Collections.Generic.IEnumerable<ModelItem> items)
        {
            foreach (ModelItem item in items)
            {
                yield return item;
                if (item.Children != null && item.Children.Any())
                {
                    foreach (var child in Traverse(item.Children))
                    {
                        yield return child;
                    }
                }
            }
        }
    }
}
