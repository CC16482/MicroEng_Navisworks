using System;
using System.Globalization; // For NumberStyles
using System.Windows;
using System.Windows.Controls;
using Autodesk.Navisworks.Api; // For ModelItem, Document, Application, ComApiBridge etc.

// Aliases for the COM Interop namespace and the ComApiBridge static class
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi; // Namespace alias
using ComApiBridge = Autodesk.Navisworks.Api.ComApiBridge; // Static class alias

// Ensure this namespace matches your project's RootNamespace (in .csproj) 
// and the x:Class attribute in your AppendDataControl.xaml file.
namespace AppendDataTool.UI
{
    public partial class AppendDataControl : UserControl, IDisposable
    {
        public AppendDataControl()
        {
            InitializeComponent();
        }

        private void AddDataButton_Click(object sender, RoutedEventArgs e)
        {
            if (StatusTextBlock == null) return; // Should not happen if XAML is correct

            StatusTextBlock.Text = "Processing...";
            try
            {
                string tabDisplayName = TabNameTextBox.Text;
                string paramDisplayName = ParamDisplayNameTextBox.Text;
                string paramInternalName = ParamInternalNameTextBox.Text;
                string paramValueString = ParamValueTextBox.Text;
                string selectedDataType = (ParamDataTypeComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();

                if (string.IsNullOrWhiteSpace(tabDisplayName) ||
                    string.IsNullOrWhiteSpace(paramDisplayName) ||
                    string.IsNullOrWhiteSpace(paramInternalName) ||
                    string.IsNullOrWhiteSpace(paramValueString) ||
                    string.IsNullOrWhiteSpace(selectedDataType))
                {
                    StatusTextBlock.Text = "Error: All fields must be filled.";
                    MessageBox.Show("All fields must be filled to proceed.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (paramInternalName.Contains(" ") || !System.Text.RegularExpressions.Regex.IsMatch(paramInternalName, @"^[a-zA-Z0-9_]+$"))
                {
                    StatusTextBlock.Text = "Error: Parameter Internal Name can only contain letters, numbers, and underscores.";
                    MessageBox.Show("Parameter Internal Name can only contain letters, numbers, and underscores.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Using Autodesk.Navisworks.Api.Application directly
                Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                if (doc == null || doc.CurrentSelection.IsEmpty)
                {
                    StatusTextBlock.Text = "Error: No document open or no items selected.";
                    MessageBox.Show("Please select one or more items in Navisworks first.", "Selection Error", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                ModelItemCollection selectedItems = doc.CurrentSelection.SelectedItems;

                // Using alias for ComApiBridge.State
                // ComApiBridge refers to Autodesk.Navisworks.Api.ComApiBridge (static class)
                // ComApi refers to Autodesk.Navisworks.Api.Interop.ComApi (namespace)
                ComApi.InwOpState10 oState = ComApiBridge.State;

                if (oState == null)
                {
                    StatusTextBlock.Text = "Error: Could not access Navisworks COM API state.";
                    MessageBox.Show("Could not access Navisworks COM API state. Please ensure Navisworks is functioning correctly.", "COM Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                int itemsProcessed = 0;
                int itemsFailed = 0;

                foreach (ModelItem oEachSelectedItem in selectedItems)
                {
                    try
                    {
                        // Using aliases for ComApiBridge.ToInwOaPath and ComApi types
                        ComApi.InwOaPath oPath = ComApiBridge.ToInwOaPath(oEachSelectedItem);
                        ComApi.InwGUIPropertyNode2 propn = (ComApi.InwGUIPropertyNode2)oState.GetGUIPropertyNode(oPath, true);
                        ComApi.InwOaPropertyVec newPvec = (ComApi.InwOaPropertyVec)oState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
                        ComApi.InwOaProperty newP = (ComApi.InwOaProperty)oState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty, null, null);

                        newP.name = paramInternalName;
                        newP.UserName = paramDisplayName;

                        object propertyValue = null;
                        bool conversionOk = false;
                        switch (selectedDataType)
                        {
                            case "Integer":
                                if (int.TryParse(paramValueString, out int intVal)) { propertyValue = intVal; conversionOk = true; }
                                break;
                            case "Double":
                                if (double.TryParse(paramValueString, NumberStyles.Any, CultureInfo.InvariantCulture, out double dblVal)) { propertyValue = dblVal; conversionOk = true; }
                                break;
                            case "Boolean (True/False)":
                                if (bool.TryParse(paramValueString, out bool boolVal)) { propertyValue = boolVal; conversionOk = true; }
                                break;
                            case "String":
                            default:
                                propertyValue = paramValueString;
                                conversionOk = true;
                                break;
                        }

                        if (!conversionOk)
                        {
                            StatusTextBlock.Text = $"Warning: Could not convert '{paramValueString}' to {selectedDataType} for item '{oEachSelectedItem.DisplayName}'. Property not added.";
                            itemsFailed++;
                            continue;
                        }
                        newP.value = propertyValue;
                        newPvec.Properties().Add(newP);

                        string tabInternalName = "UserTab_" + System.Text.RegularExpressions.Regex.Replace(tabDisplayName, @"[^A-Za-z0-9_]", "");
                        propn.SetUserDefined(0, tabInternalName, tabDisplayName, newPvec);
                        itemsProcessed++;
                    }
                    catch (Exception itemEx)
                    {
                        StatusTextBlock.Text = $"Error processing item '{oEachSelectedItem.DisplayName}': {itemEx.Message}";
                        itemsFailed++;
                    }
                }

                string summaryMessage = $"Processed {selectedItems.Count} item(s). Added/updated data for {itemsProcessed} item(s).";
                if (itemsFailed > 0)
                {
                    summaryMessage += $" Failed for {itemsFailed} item(s)/properties.";
                }
                StatusTextBlock.Text = summaryMessage;
                MessageBox.Show(summaryMessage, "Operation Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"An unexpected error occurred: {ex.Message}";
                MessageBox.Show($"An unexpected error occurred: {ex.Message}\n\n{ex.StackTrace}", "Critical Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void Dispose()
        {
            // Nothing specific to dispose in this UserControl for now
        }
    }
}