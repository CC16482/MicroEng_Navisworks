using System.Drawing;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using DrawingColor = System.Drawing.Color;
using System.Linq;

namespace MicroEng.Navisworks
{
    internal class PropertyPickResult
    {
        public string CategoryName { get; set; }
        public string CategoryDisplayName { get; set; }
        public string PropertyName { get; set; }
        public string PropertyDisplayName { get; set; }

        public string ToPath() => $"{CategoryName}|{PropertyName}";

        public string ToLabel(bool useInternalNames)
        {
            var category = useInternalNames ? CategoryName : (CategoryDisplayName ?? CategoryName);
            var prop = useInternalNames ? PropertyName : (PropertyDisplayName ?? PropertyName);
            return $"{category} > {prop}";
        }
    }

    internal class PropertyPickerDialog : Form
    {
        private readonly TreeView _tree;
        private readonly Label _itemLabel;
        private readonly ModelItem _sampleItem;
        private readonly bool _showInternalNames;

        public PropertyPickResult SelectedProperty { get; private set; }

        public PropertyPickerDialog(ModelItem sampleItem, bool showInternalNames)
        {
            _sampleItem = sampleItem;
            _showInternalNames = showInternalNames;

            Text = "Select Source Property";
            Width = 420;
            Height = 520;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = ThemeAssets.DefaultFont;
            BackColor = ThemeAssets.BackgroundPanel;

            _tree = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false,
                BackColor = DrawingColor.White,
                ForeColor = ThemeAssets.TextPrimary
            };
            _tree.NodeMouseDoubleClick += (_, _) => ConfirmSelection();

            _itemLabel = new Label
            {
                Dock = DockStyle.Top,
                Padding = new Padding(8),
                Text = _sampleItem != null ? $"Sample: {_sampleItem.DisplayName}" : "No sample item selected.",
                ForeColor = ThemeAssets.TextSecondary
            };

            var okButton = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Dock = DockStyle.Right,
                Width = 80
            };
            okButton.Click += (_, _) => ConfirmSelection();

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Dock = DockStyle.Right,
                Width = 80
            };

            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40,
                Padding = new Padding(8)
            };
            bottomPanel.Controls.Add(cancelButton);
            bottomPanel.Controls.Add(okButton);

            StylePrimaryButton(okButton);
            StyleSecondaryButton(cancelButton);

            Controls.Add(_tree);
            Controls.Add(bottomPanel);
            Controls.Add(_itemLabel);

            LoadProperties();
        }

        private void LoadProperties()
        {
            if (_sampleItem == null)
            {
                // Try to populate from cached scrape
                _tree.BeginUpdate();
                _tree.Nodes.Clear();
                var cached = DataScraperCache.LastSession?.Properties;
                if (cached != null)
                {
                    foreach (var group in cached.GroupBy(p => p.Category))
                    {
                        var catNode = new TreeNode(group.Key);
                        foreach (var prop in group.OrderBy(p => p.Name))
                        {
                            catNode.Nodes.Add(new TreeNode(prop.Name)
                            {
                                Tag = new PropertyPickResult
                                {
                                    CategoryName = prop.Category,
                                    CategoryDisplayName = prop.Category,
                                    PropertyName = prop.Name,
                                    PropertyDisplayName = prop.Name
                                }
                            });
                        }
                        _tree.Nodes.Add(catNode);
                    }
                    _tree.ExpandAll();
                }
                _tree.EndUpdate();
                return;
            }

            _tree.BeginUpdate();
            _tree.Nodes.Clear();

            foreach (var category in _sampleItem.PropertyCategories)
            {
                if (category == null) continue;
                var catNode = new TreeNode(_showInternalNames ? category.Name : category.DisplayName)
                {
                    Tag = null
                };

                foreach (var prop in category.Properties)
                {
                    var label = _showInternalNames ? prop.Name : prop.DisplayName;
                    string tooltip = null;
                    try
                    {
                        if (prop.Value != null && prop.Value.IsDisplayString)
                        {
                            tooltip = prop.Value.ToDisplayString();
                        }
                        else if (prop.Value != null)
                        {
                            tooltip = prop.Value.ToString();
                        }
                    }
                    catch
                    {
                        // ignore tooltip failures
                        tooltip = null;
                    }

                    var propNode = new TreeNode(label)
                    {
                        Tag = new PropertyPickResult
                        {
                            CategoryName = category.Name,
                            CategoryDisplayName = category.DisplayName,
                            PropertyName = prop.Name,
                            PropertyDisplayName = prop.DisplayName
                        },
                        ToolTipText = tooltip
                    };

                    catNode.Nodes.Add(propNode);
                }

                _tree.Nodes.Add(catNode);
            }

            _tree.ExpandAll();
            _tree.EndUpdate();
        }

        private void ConfirmSelection()
        {
            if (_tree.SelectedNode?.Tag is PropertyPickResult pick)
            {
                SelectedProperty = pick;
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        private static void StylePrimaryButton(Button button)
        {
            button.BackColor = ThemeAssets.Accent;
            button.ForeColor = DrawingColor.White;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = ThemeAssets.AccentStrong;
        }

        private static void StyleSecondaryButton(Button button)
        {
            button.BackColor = ThemeAssets.BackgroundMuted;
            button.ForeColor = ThemeAssets.TextPrimary;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderColor = ThemeAssets.Accent;
            button.FlatAppearance.BorderSize = 1;
        }
    }
}
