using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using Autodesk.Navisworks.Api;
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace MicroEng.Navisworks
{
    internal partial class DataMatrixColumnBuilderWindow : Window
    {
        private readonly ColumnBuilderViewModel _viewModel;
        public event EventHandler Applied;
        public event EventHandler Cancelled;
        private Point _dragStartPoint;
        private ColumnBuilderPropertyNode _lastDragTarget;
        private bool _lastDragInsertAfter;
        private bool _pendingClearSelection;
        private ColumnBuilderPropertyNode _pendingClearItem;
        private readonly DispatcherTimer _leftScrollLogTimer;
        private readonly DispatcherTimer _leftScrollIdleTimer;
        private readonly DispatcherTimer _uiHeartbeatTimer;
        private readonly DispatcherTimer _selectionRefreshTimer;
        private bool _applied;
        private bool _diagnosticsEnabled;
        private bool _leftScrollPending;
        private double _leftScrollOffset;
        private double _leftScrollDelta;
        private double _leftScrollViewport;
        private double _leftScrollExtent;
        private int _leftScrollEventsSinceLog;
        private DateTime _lastScrollEventUtc;
        private DateTime _lastScrollLogUtc;
        private DateTime _lastHeartbeatUtc;
        private bool _heartbeatArmed;
        private string _lastChooseCategoryFilterText;
        private string _lastChoosePropertyFilterText;
        private IReadOnlyCollection<string> _lastSelectionKeys;
        private CancellationTokenSource _selectionScanCts;
        private Task _selectionScanTask;
        private int _selectionScanVersion;

        public DataMatrixColumnBuilderWindow(
            IEnumerable<DataMatrixAttributeDefinition> attributes,
            IList<string> currentVisibleOrder,
            IEnumerable<string> requiredAttributeIds = null)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);
            _viewModel = new ColumnBuilderViewModel(attributes, currentVisibleOrder, requiredAttributeIds);
            DataContext = _viewModel;

            MoveSelectedUpButton.Click += (_, __) => _viewModel.MoveSelectedBy(-1);
            MoveSelectedDownButton.Click += (_, __) => _viewModel.MoveSelectedBy(1);
            DeselectSelectedButton.Click += (_, __) => _viewModel.DeselectSelected();
            ClearAllSelectedButton.Click += (_, __) => _viewModel.ClearAllSelected();

            var traceEnv = Environment.GetEnvironmentVariable("MICROENG_COLUMNBUILDER_TRACE");
            _diagnosticsEnabled =
                string.Equals(traceEnv, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(traceEnv, "true", StringComparison.OrdinalIgnoreCase);
            ColumnBuilderDiagnostics.Enabled = _diagnosticsEnabled;
            MicroEngActions.Log($"[ColumnBuilder][Perf] enabled={ColumnBuilderDiagnostics.Enabled.ToString().ToLowerInvariant()}");
            _leftScrollLogTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(180)
            };
            _leftScrollLogTimer.Tick += LeftScrollLogTimer_Tick;
            _leftScrollIdleTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(250)
            };
            _leftScrollIdleTimer.Tick += LeftScrollIdleTimer_Tick;
            _uiHeartbeatTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromSeconds(2)
            };
            _uiHeartbeatTimer.Tick += UiHeartbeatTimer_Tick;
            _selectionRefreshTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(500)
            };
            _selectionRefreshTimer.Tick += SelectionRefreshTimer_Tick;
            _viewModel.PropertyChanged += ViewModel_PropertyChanged;
            HookDiagnostics();
        }

        public IReadOnlyList<string> VisibleAttributeIds => _viewModel.VisibleAttributeIds;
        public IReadOnlyList<string> RequiredAttributeIds => _viewModel.RequiredAttributeIds;

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            _applied = true;
            Applied?.Invoke(this, EventArgs.Empty);
            Close();
        }

        private void ViewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ColumnBuilderViewModel.FilterBySelection))
            {
                if (_viewModel.FilterBySelection)
                {
                    StartSelectionWatcher();
                    UpdateSelectionFilterFromSelection();
                }
                else
                {
                    StopSelectionWatcher();
                    CancelSelectionScan();
                    _viewModel.SetSelectionFilterIds(null);
                }
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _leftScrollIdleTimer.Stop();
            StopSelectionWatcher();
            CancelSelectionScan();
            if (!_applied)
            {
                Cancelled?.Invoke(this, EventArgs.Empty);
            }
        }

        private void CategoryCheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox checkbox
                && checkbox.DataContext is ColumnBuilderCategoryNode node)
            {
                node.ApplyUserToggle();
                if (_diagnosticsEnabled)
                {
                    LogDiag($"ChooseTree toggle category '{node.Name}' -> {DescribeCheckState(node.IsChecked)}");
                }
                e.Handled = true;
            }
        }

        private void CategoryCheckBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Space && e.Key != Key.Enter)
            {
                return;
            }

            if (sender is System.Windows.Controls.CheckBox checkbox
                && checkbox.DataContext is ColumnBuilderCategoryNode node)
            {
                node.ApplyUserToggle();
                if (_diagnosticsEnabled)
                {
                    LogDiag($"ChooseTree toggle category '{node.Name}' -> {DescribeCheckState(node.IsChecked)} (key)");
                }
                e.Handled = true;
            }
        }

        private void SelectedListBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragStartPoint = e.GetPosition(null);
            _pendingClearSelection = false;
            _pendingClearItem = null;

            if (SelectedListBox == null) return;
            if (IsWithinCheckBox(e.OriginalSource as DependencyObject)) return;

            var container = ItemsControl.ContainerFromElement(SelectedListBox, e.OriginalSource as DependencyObject) as ListBoxItem;
            if (container?.IsSelected == true)
            {
                _pendingClearSelection = true;
                _pendingClearItem = container.DataContext as ColumnBuilderPropertyNode;
            }
        }

        private void SelectedListBox_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_pendingClearSelection) return;

            _pendingClearSelection = false;
            if (SelectedListBox == null) return;

            if (_pendingClearItem != null && Equals(SelectedListBox.SelectedItem, _pendingClearItem))
            {
                SelectedListBox.SelectedItem = null;
                _viewModel.SelectedItem = null;
                e.Handled = true;
            }
        }

        private void SelectedListBox_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed) return;

            var position = e.GetPosition(null);
            var dx = Math.Abs(position.X - _dragStartPoint.X);
            var dy = Math.Abs(position.Y - _dragStartPoint.Y);
            if (dx < SystemParameters.MinimumHorizontalDragDistance
                && dy < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            _pendingClearSelection = false;
            if (SelectedListBox?.SelectedItem is ColumnBuilderPropertyNode node)
            {
                _lastDragTarget = null;
                _lastDragInsertAfter = false;
                DragDrop.DoDragDrop(SelectedListBox, new DataObject(typeof(ColumnBuilderPropertyNode), node), DragDropEffects.Move);
            }
        }

        private void SelectedListBox_DragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(ColumnBuilderPropertyNode)))
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            e.Effects = DragDropEffects.Move;
            e.Handled = true;

            var source = e.Data.GetData(typeof(ColumnBuilderPropertyNode)) as ColumnBuilderPropertyNode;
            if (source == null) return;

            if (!_viewModel.CanReorderSelected)
            {
                return;
            }

            var target = GetDropTarget(e, out var insertAfter);
            if (target == null || target == source) return;

            if (ReferenceEquals(target, _lastDragTarget) && insertAfter == _lastDragInsertAfter)
            {
                return;
            }

            if (_viewModel.MoveSelectedTo(source, target, insertAfter))
            {
                _lastDragTarget = target;
                _lastDragInsertAfter = insertAfter;
            }
        }

        private void SelectedListBox_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(ColumnBuilderPropertyNode))) return;

            var source = e.Data.GetData(typeof(ColumnBuilderPropertyNode)) as ColumnBuilderPropertyNode;
            if (source == null) return;

            var target = GetDropTarget(e, out var insertAfter);
            _viewModel.MoveSelectedTo(source, target, insertAfter);
        }

        private ColumnBuilderPropertyNode GetDropTarget(DragEventArgs e, out bool insertAfter)
        {
            insertAfter = false;
            if (SelectedListBox == null) return null;

            var source = e.OriginalSource as DependencyObject;
            var container = ItemsControl.ContainerFromElement(SelectedListBox, source) as ListBoxItem;
            if (container != null)
            {
                var pos = e.GetPosition(container);
                insertAfter = pos.Y > container.ActualHeight / 2.0;
                return container.DataContext as ColumnBuilderPropertyNode;
            }

            if (SelectedListBox.Items.Count == 0) return null;

            var listPos = e.GetPosition(SelectedListBox);
            var lastIndex = SelectedListBox.Items.Count - 1;
            var lastContainer = SelectedListBox.ItemContainerGenerator.ContainerFromIndex(lastIndex) as ListBoxItem;
            if (lastContainer == null) return null;

            var lastBottom = lastContainer.TranslatePoint(
                new Point(0, lastContainer.ActualHeight),
                SelectedListBox).Y;

            if (listPos.Y > lastBottom)
            {
                insertAfter = true;
                return lastContainer.DataContext as ColumnBuilderPropertyNode;
            }

            return null;
        }

        private static bool IsWithinCheckBox(DependencyObject source)
        {
            var current = source;
            while (current != null)
            {
                if (current is ToggleButton) return true;
                current = VisualTreeHelper.GetParent(current);
            }
            return false;
        }

        private void HookDiagnostics()
        {
            try
            {
                if (ChooseTree != null)
                {
                    ChooseTree.AddHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler(ChooseTree_ScrollChanged), true);
                }

                if (!_diagnosticsEnabled)
                {
                    return;
                }

                LogDiag("ColumnBuilder left-side diagnostics enabled.");

                if (ChooseTree != null)
                {
                    ChooseTree.AddHandler(TreeViewItem.ExpandedEvent, new RoutedEventHandler(ChooseTree_ItemExpanded), true);
                    ChooseTree.AddHandler(TreeViewItem.CollapsedEvent, new RoutedEventHandler(ChooseTree_ItemCollapsed), true);
                    ChooseTree.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(ChooseTree_CheckBoxClick), true);
                    ChooseTree.SelectedItemChanged += ChooseTree_SelectedItemChanged;
                }

                if (ChooseCategoryFilterTextBox != null)
                {
                    ChooseCategoryFilterTextBox.TextChanged += ChooseFilterTextBox_TextChanged;
                }

                if (ChoosePropertyFilterTextBox != null)
                {
                    ChoosePropertyFilterTextBox.TextChanged += ChooseFilterTextBox_TextChanged;
                }

                _lastHeartbeatUtc = DateTime.UtcNow;
                _heartbeatArmed = true;
                _uiHeartbeatTimer.Start();
            }
            catch (Exception ex)
            {
                if (_diagnosticsEnabled)
                {
                    LogDiag($"ColumnBuilder diagnostics hook failed: {ex}");
                }
            }
        }

        private void ChooseFilterTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SyncFilterTextFromUi();
            if (!_diagnosticsEnabled)
            {
                return;
            }

            try
            {
                var categoryText = ChooseCategoryFilterTextBox?.Text ?? string.Empty;
                var propertyText = ChoosePropertyFilterTextBox?.Text ?? string.Empty;
                if (string.Equals(categoryText, _lastChooseCategoryFilterText, StringComparison.Ordinal)
                    && string.Equals(propertyText, _lastChoosePropertyFilterText, StringComparison.Ordinal))
                {
                    return;
                }

                _lastChooseCategoryFilterText = categoryText;
                _lastChoosePropertyFilterText = propertyText;
                LogDiag($"ChooseTree filter category='{categoryText}' property='{propertyText}'");
            }
            catch (Exception ex)
            {
                LogDiag($"ChooseTree filter log failed: {ex.Message}");
            }
        }

        private void SyncFilterTextFromUi()
        {
            if (_viewModel == null)
            {
                return;
            }

            var categoryText = ChooseCategoryFilterTextBox?.Text ?? string.Empty;
            var propertyText = ChoosePropertyFilterTextBox?.Text ?? string.Empty;

            _viewModel.CategoryFilterText = categoryText;
            _viewModel.PropertyFilterText = propertyText;
        }

        private void ChooseTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (!_diagnosticsEnabled)
            {
                return;
            }

            try
            {
                LogDiag($"ChooseTree select {DescribeNode(e.NewValue)}");
            }
            catch (Exception ex)
            {
                LogDiag($"ChooseTree select log failed: {ex.Message}");
            }
        }

        private void ChooseTree_ItemExpanded(object sender, RoutedEventArgs e)
        {
            if (!_diagnosticsEnabled)
            {
                return;
            }

            try
            {
                if (e.OriginalSource is TreeViewItem item)
                {
                    LogDiag($"ChooseTree expand {DescribeNode(item.DataContext)}");
                }
            }
            catch (Exception ex)
            {
                LogDiag($"ChooseTree expand log failed: {ex.Message}");
            }
        }

        private void ChooseTree_ItemCollapsed(object sender, RoutedEventArgs e)
        {
            if (!_diagnosticsEnabled)
            {
                return;
            }

            try
            {
                if (e.OriginalSource is TreeViewItem item)
                {
                    LogDiag($"ChooseTree collapse {DescribeNode(item.DataContext)}");
                }
            }
            catch (Exception ex)
            {
                LogDiag($"ChooseTree collapse log failed: {ex.Message}");
            }
        }

        private void ChooseTreeItem_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is TreeViewItem item)
            {
                AnimateTreeExpander(item, item.IsExpanded, animate: false);
            }
        }

        private void ChooseTreeItem_Expanded(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is TreeViewItem item)
            {
                AnimateTreeExpander(item, true, animate: true);
                if (_viewModel.FilterBySelection && item.DataContext is ColumnBuilderCategoryNode category)
                {
                    category.RefreshPropertiesView();
                }
            }
        }

        private void ChooseTreeItem_Collapsed(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is TreeViewItem item)
            {
                AnimateTreeExpander(item, false, animate: true);
            }
        }

        private void AnimateTreeExpander(TreeViewItem item, bool isExpanded, bool animate)
        {
            if (item == null) return;

            var arrowContainer = FindVisualChildByName<FrameworkElement>(item, "ArrowContainer");
            if (arrowContainer == null)
            {
                item.Dispatcher.BeginInvoke(new Action(() =>
                    AnimateTreeExpander(item, isExpanded, animate)), DispatcherPriority.Loaded);
                return;
            }

            var rotate = arrowContainer.RenderTransform as RotateTransform;
            if (rotate == null || rotate.IsFrozen)
            {
                rotate = new RotateTransform();
                arrowContainer.RenderTransform = rotate;
                arrowContainer.RenderTransformOrigin = new Point(0.5, 0.5);
            }

            var targetAngle = isExpanded ? 0d : -90d;
            if (!animate)
            {
                rotate.Angle = targetAngle;
                return;
            }

            var animation = new DoubleAnimation(targetAngle, TimeSpan.FromMilliseconds(150))
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
            };
            rotate.BeginAnimation(RotateTransform.AngleProperty, animation);
        }

        private static T FindVisualChildByName<T>(DependencyObject root, string name) where T : FrameworkElement
        {
            if (root == null) return null;

            var count = VisualTreeHelper.GetChildrenCount(root);
            for (var i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(root, i);
                if (child is T element && element.Name == name)
                {
                    return element;
                }

                var found = FindVisualChildByName<T>(child, name);
                if (found != null)
                {
                    return found;
                }
            }

            return null;
        }

        private void ChooseTree_CheckBoxClick(object sender, RoutedEventArgs e)
        {
            if (!_diagnosticsEnabled)
            {
                return;
            }

            try
            {
                if (e.OriginalSource is CheckBox cb)
                {
                    if (cb.DataContext is ColumnBuilderPropertyNode prop)
                    {
                        LogDiag($"ChooseTree toggle property '{prop.CategoryName}.{prop.PropertyName}' -> {DescribeCheckState(cb.IsChecked)}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogDiag($"ChooseTree toggle log failed: {ex.Message}");
            }
        }

        private void ChooseTree_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            var hasDelta = Math.Abs(e.VerticalChange) > 0.1 || Math.Abs(e.HorizontalChange) > 0.1;
            if (!hasDelta)
            {
                return;
            }

            _viewModel.SetTreeScrolling(true);
            _leftScrollIdleTimer.Stop();
            _leftScrollIdleTimer.Start();

            if (!_diagnosticsEnabled)
            {
                return;
            }

            try
            {
                _leftScrollOffset = e.VerticalOffset;
                _leftScrollDelta = e.VerticalChange;
                _leftScrollViewport = e.ViewportHeight;
                _leftScrollExtent = e.ExtentHeight;
                _leftScrollEventsSinceLog++;
                _leftScrollPending = true;
                _lastScrollEventUtc = DateTime.UtcNow;

                if (!_leftScrollLogTimer.IsEnabled)
                {
                    _lastScrollLogUtc = DateTime.MinValue;
                    _leftScrollLogTimer.Start();
                }
            }
            catch (Exception ex)
            {
                LogDiag($"ChooseTree scroll log failed: {ex.Message}");
            }
        }

        private void LeftScrollLogTimer_Tick(object sender, EventArgs e)
        {
            if (!_diagnosticsEnabled)
            {
                _leftScrollLogTimer.Stop();
                return;
            }

            if (!_leftScrollPending)
            {
                _leftScrollLogTimer.Stop();
                return;
            }

            var now = DateTime.UtcNow;
            var sinceEventMs = (now - _lastScrollEventUtc).TotalMilliseconds;

            if (sinceEventMs > 350)
            {
                LogDiag($"ChooseTree scroll stop offset={_leftScrollOffset:F1} viewport={_leftScrollViewport:F1} extent={_leftScrollExtent:F1}");
                _leftScrollEventsSinceLog = 0;
                _leftScrollPending = false;
                _leftScrollLogTimer.Stop();
                return;
            }

            if ((now - _lastScrollLogUtc).TotalMilliseconds >= 250)
            {
                LogDiag($"ChooseTree scroll offset={_leftScrollOffset:F1} Δ={_leftScrollDelta:F1} viewport={_leftScrollViewport:F1} extent={_leftScrollExtent:F1} events=+{_leftScrollEventsSinceLog}");
                _leftScrollEventsSinceLog = 0;
                _lastScrollLogUtc = now;
            }
        }

        private void LeftScrollIdleTimer_Tick(object sender, EventArgs e)
        {
            _leftScrollIdleTimer.Stop();
            _viewModel.SetTreeScrolling(false);
        }

        private void ChooseTreeItem_Selected(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is TreeViewItem item)
            {
                item.IsSelected = false;
                e.Handled = true;
            }
        }

        private void UiHeartbeatTimer_Tick(object sender, EventArgs e)
        {
            if (!_diagnosticsEnabled)
            {
                _uiHeartbeatTimer.Stop();
                return;
            }

            var now = DateTime.UtcNow;
            if (_heartbeatArmed)
            {
                var gap = now - _lastHeartbeatUtc;
                if (gap > TimeSpan.FromSeconds(5))
                {
                    LogDiag($"UI stall detected ({gap.TotalSeconds:F1}s since last heartbeat)");
                }
            }

            _lastHeartbeatUtc = now;
            _heartbeatArmed = true;
        }

        private static string DescribeNode(object node)
        {
            return node switch
            {
                ColumnBuilderCategoryNode category => $"category '{category.Name}'",
                ColumnBuilderPropertyNode property => $"property '{property.CategoryName}.{property.PropertyName}'",
                null => "<null>",
                _ => node.GetType().Name
            };
        }

        private static string DescribeCheckState(bool? state)
        {
            if (state == true) return "checked";
            if (state == false) return "unchecked";
            return "indeterminate";
        }

        private static void LogDiag(string message)
        {
            MicroEngActions.Log($"[ColumnBuilder][Choose] {message}");
        }

        private void StartSelectionWatcher()
        {
            _lastSelectionKeys = null;
            if (!_selectionRefreshTimer.IsEnabled)
            {
                _selectionRefreshTimer.Start();
            }
        }

        private void StopSelectionWatcher()
        {
            if (_selectionRefreshTimer.IsEnabled)
            {
                _selectionRefreshTimer.Stop();
            }
            _lastSelectionKeys = null;
        }

        private void CancelSelectionScan()
        {
            try
            {
                _selectionScanCts?.Cancel();
            }
            catch
            {
                // ignore
            }

            _selectionScanCts?.Dispose();
            _selectionScanCts = null;
            _selectionScanTask = null;
        }

        private void SelectionRefreshTimer_Tick(object sender, EventArgs e)
        {
            if (!_viewModel.FilterBySelection)
            {
                StopSelectionWatcher();
                return;
            }

            var keys = GetCurrentSelectionKeys();
            if (SelectionKeysEqual(_lastSelectionKeys, keys))
            {
                return;
            }

            _lastSelectionKeys = keys;
            UpdateSelectionFilterFromSelection();
        }

        private void UpdateSelectionFilterFromSelection()
        {
            if (!_viewModel.FilterBySelection)
            {
                return;
            }

            var session = DataScraperCache.LastSession;
            var doc = NavisApp.ActiveDocument;
            if (session == null || doc?.CurrentSelection?.SelectedItems == null)
            {
                _viewModel.SetSelectionFilterIds(new HashSet<string>(StringComparer.OrdinalIgnoreCase));
                return;
            }

            var sw = ColumnBuilderDiagnostics.Start("SelectionFilter");
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ModelItem item in doc.CurrentSelection.SelectedItems)
            {
                try
                {
                    keys.Add(item.InstanceGuid.ToString("D"));
                }
                catch
                {
                    // ignore
                }
            }

            if (keys.Count == 0)
            {
                _viewModel.SetSelectionFilterIds(new HashSet<string>(StringComparer.OrdinalIgnoreCase));
                ColumnBuilderDiagnostics.End(sw, "SelectionFilter", $"keys=0");
                return;
            }

            StartSelectionScan(session, keys, sw);
        }

        private void StartSelectionScan(ScrapeSession session, HashSet<string> keys, Stopwatch sw)
        {
            CancelSelectionScan();

            _selectionScanCts = new CancellationTokenSource();
            var token = _selectionScanCts.Token;
            var scanVersion = ++_selectionScanVersion;
            var keysSnapshot = new HashSet<string>(keys, StringComparer.OrdinalIgnoreCase);

            ColumnBuilderDiagnostics.End(sw, "SelectionFilter", $"keys={keysSnapshot.Count} scan=start");

            _selectionScanTask = Task.Run(() =>
            {
                return BuildSelectionFilterIdsForKeys(session, keysSnapshot, token);
            }, token).ContinueWith(t =>
            {
                if (t.IsCanceled || t.IsFaulted) return;
                var ids = t.Result ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (!_viewModel.FilterBySelection) return;
                    if (scanVersion != _selectionScanVersion) return;
                    _viewModel.SetSelectionFilterIds(ids);
                    ColumnBuilderDiagnostics.Log($"SelectionFilter scan applied ids={ids.Count}");
                }), DispatcherPriority.Background);
            }, CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.Default);
        }

        private static HashSet<string> BuildSelectionFilterIdsForKeys(ScrapeSession session, HashSet<string> keys, CancellationToken token)
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (session == null || keys == null || keys.Count == 0)
            {
                return ids;
            }

            var rawEntries = session.RawEntries;
            if (rawEntries == null || rawEntries.Count == 0)
            {
                return ids;
            }

            var idCache = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in rawEntries)
            {
                if (token.IsCancellationRequested)
                {
                    return ids;
                }

                if (!keys.Contains(entry.ItemKey))
                {
                    continue;
                }

                var category = entry.Category ?? string.Empty;
                var name = entry.Name ?? string.Empty;

                if (!idCache.TryGetValue(category, out var byName))
                {
                    byName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    idCache[category] = byName;
                }

                if (!byName.TryGetValue(name, out var id))
                {
                    id = string.Concat(category, "|", name);
                    byName[name] = id;
                }

                ids.Add(id);
            }

            return ids;
        }

        private static IReadOnlyCollection<string> GetCurrentSelectionKeys()
        {
            var doc = NavisApp.ActiveDocument;
            if (doc?.CurrentSelection?.SelectedItems == null)
            {
                return Array.Empty<string>();
            }

            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ModelItem item in doc.CurrentSelection.SelectedItems)
            {
                try
                {
                    keys.Add(item.InstanceGuid.ToString("D"));
                }
                catch
                {
                    // ignore
                }
            }
            return keys;
        }

        private static bool SelectionKeysEqual(IReadOnlyCollection<string> a, IReadOnlyCollection<string> b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null) return false;
            if (a.Count != b.Count) return false;
            if (a.Count == 0) return true;

            var set = new HashSet<string>(a, StringComparer.OrdinalIgnoreCase);
            foreach (var key in b)
            {
                if (!set.Contains(key))
                {
                    return false;
                }
            }
            return true;
        }
    }

    internal sealed class ColumnBuilderViewModel : INotifyPropertyChanged
    {
        private readonly List<ColumnBuilderPropertyNode> _allProperties = new();
        private ObservableCollection<ColumnBuilderPropertyNode> _selectedProperties;
        private ObservableCollection<ColumnBuilderPropertyNode> _filteredProperties;
        private List<ColumnBuilderPropertyNode> _selectedAll = new();
        private bool _suppressSelectionChanged;
        private string _categoryFilterText;
        private string _propertyFilterText;
        private string _selectedFilterText;
        private string _selectedCountText = "0 selected";
        private int _selectedCountValue;
        private string _selectedCountSuffix = " selected";
        private int _availableCount;
        private ColumnBuilderPropertyNode _selectedItem;
        private readonly DispatcherTimer _rebuildTimer;
        private readonly DispatcherTimer _filterTimer;
        private bool _isFilterActive;
        private bool _filterBySelection;
        private HashSet<string> _selectionFilteredIds;
        private HashSet<string> _selectionFilteredCategories;
        private bool _selectionEmpty;
        private bool _isChooseTreeScrolling;

        public ColumnBuilderViewModel(
            IEnumerable<DataMatrixAttributeDefinition> attributes,
            IList<string> currentVisibleOrder,
            IEnumerable<string> requiredAttributeIds)
        {
            Categories = new ObservableCollection<ColumnBuilderCategoryNode>();
            SelectedProperties = new ObservableCollection<ColumnBuilderPropertyNode>();
            FilteredProperties = new ObservableCollection<ColumnBuilderPropertyNode>();
            _rebuildTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(120)
            };
            _rebuildTimer.Tick += (_, __) =>
            {
                _rebuildTimer.Stop();
                RebuildSelected();
            };
            _filterTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(160)
            };
            _filterTimer.Tick += (_, __) =>
            {
                _filterTimer.Stop();
                ApplyFilter();
            };

            var orderMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (currentVisibleOrder != null)
            {
                for (var i = 0; i < currentVisibleOrder.Count; i++)
                {
                    var id = currentVisibleOrder[i];
                    if (string.IsNullOrWhiteSpace(id)) continue;
                    if (!orderMap.ContainsKey(id))
                    {
                        orderMap[id] = i;
                    }
                }
            }

            var available = attributes?.ToList() ?? new List<DataMatrixAttributeDefinition>();
            var requiredSet = new HashSet<string>(
                requiredAttributeIds ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);
            var byCategory = available
                .GroupBy(a => a.Category ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var group in byCategory)
            {
                var category = new ColumnBuilderCategoryNode(group.Key, IsPropertyVisible);
                category.SelectionChanged += OnSelectionChanged;

                foreach (var attr in group.OrderBy(a => a.PropertyName, StringComparer.OrdinalIgnoreCase))
                {
                    var id = attr.Id ?? $"{attr.Category}|{attr.PropertyName}";
                    var propName = string.IsNullOrWhiteSpace(attr.DisplayName) ? attr.PropertyName : attr.DisplayName;
                    if (string.IsNullOrWhiteSpace(propName)) propName = id;

                    var node = new ColumnBuilderPropertyNode(id, propName, group.Key)
                    {
                        SortIndex = orderMap.TryGetValue(id, out var index) ? index : int.MaxValue,
                        IsChecked = orderMap.ContainsKey(id)
                    };
                    if (requiredSet.Contains(id))
                    {
                        node.IsRequired = true;
                        if (node.IsChecked != true)
                        {
                            node.IsChecked = true;
                        }
                    }
                    node.SelectionChanged += OnSelectionChanged;
                    node.Parent = category;
                    category.Properties.Add(node);
                    _allProperties.Add(node);
                }

                category.IsExpanded = category.Properties.Any(p => p.IsChecked == true);
                category.RefreshFromChildren();
                Categories.Add(category);
            }

            RebuildSelected();
            ApplyFilter();
        }

        public ObservableCollection<ColumnBuilderCategoryNode> Categories { get; }

        public ObservableCollection<ColumnBuilderPropertyNode> SelectedProperties
        {
            get => _selectedProperties;
            private set
            {
                if (_selectedProperties == value) return;
                _selectedProperties = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<ColumnBuilderPropertyNode> FilteredProperties
        {
            get => _filteredProperties;
            private set
            {
                if (_filteredProperties == value) return;
                _filteredProperties = value;
                OnPropertyChanged();
            }
        }

        public bool IsFilterActive
        {
            get => _isFilterActive;
            private set
            {
                if (_isFilterActive == value) return;
                _isFilterActive = value;
                OnPropertyChanged();
            }
        }

        public IReadOnlyList<string> VisibleAttributeIds => SelectedProperties.Select(p => p.Id).ToList();
        public IReadOnlyList<string> RequiredAttributeIds =>
            _allProperties
                .Where(p => p.IsChecked == true && p.IsRequired)
                .Select(p => p.Id)
                .ToList();

        public string SelectedCountText
        {
            get => _selectedCountText;
            private set
            {
                if (_selectedCountText == value) return;
                _selectedCountText = value;
                OnPropertyChanged();
            }
        }

        public int SelectedCountValue
        {
            get => _selectedCountValue;
            private set
            {
                if (_selectedCountValue == value) return;
                _selectedCountValue = value;
                OnPropertyChanged();
            }
        }

        public string SelectedCountSuffix
        {
            get => _selectedCountSuffix;
            private set
            {
                if (_selectedCountSuffix == value) return;
                _selectedCountSuffix = value;
                OnPropertyChanged();
            }
        }

        public int AvailableCount
        {
            get => _availableCount;
            private set
            {
                if (_availableCount == value) return;
                _availableCount = value;
                OnPropertyChanged();
            }
        }

        public string CategoryFilterText
        {
            get => _categoryFilterText;
            set
            {
                if (_categoryFilterText == value) return;
                _categoryFilterText = value;
                OnPropertyChanged();
                ScheduleFilter();
            }
        }

        public string PropertyFilterText
        {
            get => _propertyFilterText;
            set
            {
                if (_propertyFilterText == value) return;
                _propertyFilterText = value;
                OnPropertyChanged();
                ScheduleFilter();
            }
        }

        public string SelectedFilterText
        {
            get => _selectedFilterText;
            set
            {
                if (_selectedFilterText == value) return;
                _selectedFilterText = value;
                OnPropertyChanged();
                ApplySelectedFilter();
            }
        }

        public ColumnBuilderPropertyNode SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (_selectedItem == value) return;
                _selectedItem = value;
                OnPropertyChanged();
            }
        }

        public bool CanReorderSelected => string.IsNullOrWhiteSpace(_selectedFilterText);

        public bool FilterBySelection
        {
            get => _filterBySelection;
            set
            {
                if (_filterBySelection == value) return;
                _filterBySelection = value;
                OnPropertyChanged();
                if (!_filterBySelection)
                {
                    _selectionFilteredIds = null;
                }
                ScheduleFilter();
            }
        }

        public bool IsChooseTreeScrolling
        {
            get => _isChooseTreeScrolling;
            private set
            {
                if (_isChooseTreeScrolling == value) return;
                _isChooseTreeScrolling = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnSelectionChanged()
        {
            if (_suppressSelectionChanged) return;
            ScheduleRebuildSelected();
        }

        private void ScheduleRebuildSelected()
        {
            _rebuildTimer.Stop();
            _rebuildTimer.Start();
        }

        private void ScheduleFilter()
        {
            _filterTimer.Stop();
            _filterTimer.Start();
        }

        private void RebuildSelected()
        {
            _suppressSelectionChanged = true;
            var sw = ColumnBuilderDiagnostics.Start("RebuildSelected");
            try
            {
                var selected = _allProperties
                    .Where(p => p.IsChecked == true)
                    .ToList();

                if (_selectedAll == null || _selectedAll.Count == 0)
                {
                    _selectedAll = selected
                        .OrderBy(p => p.SortIndex)
                        .ToList();
                }
                else
                {
                    var selectedIds = new HashSet<string>(
                        selected.Select(p => p.Id),
                        StringComparer.OrdinalIgnoreCase);

                    _selectedAll = _selectedAll
                        .Where(p => selectedIds.Contains(p.Id))
                        .ToList();

                    var existingIds = new HashSet<string>(
                        _selectedAll.Select(p => p.Id),
                        StringComparer.OrdinalIgnoreCase);

                    var newOnes = selected
                        .Where(p => !existingIds.Contains(p.Id))
                        .OrderBy(p => p.SortIndex)
                        .ToList();

                    if (newOnes.Count > 0)
                    {
                        _selectedAll.AddRange(newOnes);
                    }
                }

                UpdateOrderIndices(updateSortIndex: false);
                ApplySelectedFilter();
            }
            finally
            {
                _suppressSelectionChanged = false;
                ColumnBuilderDiagnostics.End(sw, "RebuildSelected", $"selected={_selectedAll.Count}");
            }
        }

        public void MoveSelectedBy(int delta)
        {
            if (SelectedItem == null) return;
            var currentIndex = _selectedAll.IndexOf(SelectedItem);
            if (currentIndex < 0) return;

            var newIndex = currentIndex + delta;
            if (newIndex < 0 || newIndex >= _selectedAll.Count) return;

            _selectedAll.RemoveAt(currentIndex);
            _selectedAll.Insert(newIndex, SelectedItem);
            UpdateOrderIndices(updateSortIndex: true);
            ApplySelectedFilter();
            SelectedItem = _selectedAll[newIndex];
        }

        public bool MoveSelectedTo(ColumnBuilderPropertyNode item, ColumnBuilderPropertyNode target, bool insertAfter)
        {
            if (item == null) return false;

            var oldIndex = _selectedAll.IndexOf(item);
            if (oldIndex < 0) return false;

            int newIndex;
            if (target == null)
            {
                newIndex = _selectedAll.Count - 1;
            }
            else
            {
                var targetIndex = _selectedAll.IndexOf(target);
                if (targetIndex < 0) return false;
                newIndex = insertAfter ? targetIndex + 1 : targetIndex;
            }

            if (oldIndex == newIndex) return false;

            _selectedAll.RemoveAt(oldIndex);
            if (newIndex > oldIndex) newIndex--;
            if (newIndex < 0) newIndex = 0;
            if (newIndex > _selectedAll.Count) newIndex = _selectedAll.Count;
            _selectedAll.Insert(newIndex, item);

            UpdateOrderIndices(updateSortIndex: true);
            ApplySelectedFilter();
            SelectedItem = item;
            return true;
        }

        public void DeselectSelected()
        {
            SelectedItem = null;
        }

        public void ClearAllSelected()
        {
            if (_selectedAll.Count == 0) return;

            _suppressSelectionChanged = true;
            try
            {
                foreach (var item in _selectedAll)
                {
                    item.SetCheckedFromParent(false);
                }
                _selectedAll.Clear();
                foreach (var category in Categories)
                {
                    category.RefreshFromChildren();
                }
                SelectedItem = null;
                ApplySelectedFilter();
            }
            finally
            {
                _suppressSelectionChanged = false;
            }
        }

        private void ApplySelectedFilter()
        {
            var sw = ColumnBuilderDiagnostics.Start("ApplySelectedFilter");
            var text = _selectedFilterText?.Trim();
            IEnumerable<ColumnBuilderPropertyNode> filtered = _selectedAll;
            if (!string.IsNullOrWhiteSpace(text))
            {
                filtered = filtered.Where(p =>
                    p.PropertyName.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0
                    || p.CategoryName.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0
                    || p.Id.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            var list = filtered.ToList();
            SelectedProperties = new ObservableCollection<ColumnBuilderPropertyNode>(list);

            SelectedCountText = string.IsNullOrWhiteSpace(text)
                ? $"{_selectedAll.Count} selected"
                : $"{list.Count} selected (of {_selectedAll.Count})";
            SelectedCountValue = string.IsNullOrWhiteSpace(text)
                ? _selectedAll.Count
                : list.Count;
            SelectedCountSuffix = string.IsNullOrWhiteSpace(text)
                ? " selected"
                : $" selected (of {_selectedAll.Count})";
            ColumnBuilderDiagnostics.End(sw, "ApplySelectedFilter", $"count={list.Count}");
        }

        private void UpdateOrderIndices(bool updateSortIndex)
        {
            for (var i = 0; i < _selectedAll.Count; i++)
            {
                if (updateSortIndex)
                {
                    _selectedAll[i].SortIndex = i;
                }
                _selectedAll[i].OrderIndex = i + 1;
            }
        }

        private void ApplyFilter()
        {
            var sw = ColumnBuilderDiagnostics.Start("ApplyFilter");
            var categoryText = _categoryFilterText?.Trim();
            var propertyText = _propertyFilterText?.Trim();
            var hasCategory = !string.IsNullOrWhiteSpace(categoryText);
            var hasProperty = !string.IsNullOrWhiteSpace(propertyText);

            if (_filterBySelection && (_selectionFilteredIds == null || _selectionFilteredIds.Count == 0))
            {
                _selectionEmpty = true;
                if (!IsFilterActive)
                {
                    IsFilterActive = true;
                }
                if (_filteredProperties.Count > 0)
                {
                    FilteredProperties = new ObservableCollection<ColumnBuilderPropertyNode>();
                }
                AvailableCount = 0;
                ColumnBuilderDiagnostics.End(sw, "ApplyFilter", "selection empty");
                return;
            }

            _selectionEmpty = false;
            var baseList = _allProperties.AsEnumerable();
            if (_filterBySelection && _selectionFilteredIds != null)
            {
                baseList = baseList.Where(p => _selectionFilteredIds.Contains(p.Id));
            }

            if (!hasCategory && !hasProperty)
            {
                if (IsFilterActive)
                {
                    IsFilterActive = false;
                }
                if (_filteredProperties.Count > 0)
                {
                    FilteredProperties = new ObservableCollection<ColumnBuilderPropertyNode>();
                }

                AvailableCount = GetAvailableCount();

                if (_filterBySelection && _selectionFilteredCategories != null)
                {
                    foreach (var category in Categories)
                    {
                        var shouldShow = _selectionFilteredCategories.Contains(category.Name);
                        if (!shouldShow)
                        {
                            if (category.IsVisible)
                            {
                                category.IsVisible = false;
                            }
                            continue;
                        }

                        if (!category.IsVisible)
                        {
                            category.IsVisible = true;
                        }

                        if (category.IsExpanded)
                        {
                            category.RefreshPropertiesView();
                        }
                    }
                    ColumnBuilderDiagnostics.End(sw, "ApplyFilter", "selection tree (expanded only)");
                }
                else
                {
                    foreach (var category in Categories)
                    {
                        category.RefreshVisibility();
                    }
                    ColumnBuilderDiagnostics.End(sw, "ApplyFilter", "no text");
                }
                return;
}

            if (!IsFilterActive)
            {
                IsFilterActive = true;
            }

            var filtered = baseList
                .Where(p =>
                {
                    var categoryMatch = !hasCategory
                        || p.CategoryName.IndexOf(categoryText, StringComparison.OrdinalIgnoreCase) >= 0;
                    var propertyMatch = !hasProperty
                        || p.PropertyName.IndexOf(propertyText, StringComparison.OrdinalIgnoreCase) >= 0
                        || p.Id.IndexOf(propertyText, StringComparison.OrdinalIgnoreCase) >= 0;
                    return categoryMatch && propertyMatch;
                })
                .ToList();

            FilteredProperties = new ObservableCollection<ColumnBuilderPropertyNode>(filtered);
            AvailableCount = filtered.Count;

            // Skip TreeView visibility changes while filtering to avoid heavy layout churn.
            ColumnBuilderDiagnostics.End(sw, "ApplyFilter", $"filtered={filtered.Count}");
        }

        private int GetAvailableCount()
        {
            if (_filterBySelection)
            {
                if (_selectionFilteredIds == null || _selectionFilteredIds.Count == 0)
                {
                    return 0;
                }

                var count = 0;
                foreach (var prop in _allProperties)
                {
                    if (_selectionFilteredIds.Contains(prop.Id))
                    {
                        count++;
                    }
                }

                return count;
            }

            return _allProperties.Count;
        }

        public void SetSelectionFilterIds(HashSet<string> ids)
        {
            _selectionFilteredIds = ids;
            _selectionFilteredCategories = null;
            _selectionEmpty = _filterBySelection && (ids == null || ids.Count == 0);
            if (_selectionEmpty)
            {
                if (!IsFilterActive)
                {
                    IsFilterActive = true;
                }
                if (_filteredProperties.Count > 0)
                {
                    FilteredProperties = new ObservableCollection<ColumnBuilderPropertyNode>();
                }
                AvailableCount = 0;
                return;
            }

            if (_filterBySelection && ids != null && ids.Count > 0)
            {
                var categories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var id in ids)
                {
                    if (string.IsNullOrWhiteSpace(id)) continue;
                    var separator = id.IndexOf('|');
                    var category = separator >= 0 ? id.Substring(0, separator) : id;
                    if (!string.IsNullOrWhiteSpace(category))
                    {
                        categories.Add(category);
                    }
                }
                _selectionFilteredCategories = categories;
            }

            ScheduleFilter();
        }

        internal bool IsPropertyVisible(ColumnBuilderPropertyNode node)
        {
            if (node == null) return false;
            if (_filterBySelection)
            {
                return _selectionFilteredIds != null && _selectionFilteredIds.Contains(node.Id);
            }

            return true;
        }

        public void SetTreeScrolling(bool isScrolling)
        {
            IsChooseTreeScrolling = isScrolling;
        }

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

    }

    internal static class ColumnBuilderDiagnostics
    {
        private const int SlowThresholdMs = 120;

        public static bool Enabled { get; set; }

        public static Stopwatch Start(string name)
        {
            if (!Enabled) return null;
            return Stopwatch.StartNew();
        }

        public static void End(Stopwatch sw, string name, string details = null)
        {
            if (sw == null || !Enabled) return;

            sw.Stop();
            if (sw.ElapsedMilliseconds < SlowThresholdMs)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(details))
            {
                MicroEngActions.Log($"[ColumnBuilder][Perf] {name} {sw.ElapsedMilliseconds}ms");
            }
            else
            {
                MicroEngActions.Log($"[ColumnBuilder][Perf] {name} {sw.ElapsedMilliseconds}ms ({details})");
            }
        }

        public static void Log(string message)
        {
            if (!Enabled) return;
            MicroEngActions.Log($"[ColumnBuilder][Perf] {message}");
        }
    }

    public sealed class ColumnBuilderCategoryNode : INotifyPropertyChanged
    {
        private readonly Func<ColumnBuilderPropertyNode, bool> _filterPredicate;
        private bool? _isChecked = false;
        private bool _suppressChildren;
        private bool _isVisible = true;
        private bool _isExpanded;
        private readonly ICollectionView _propertiesView;

        public ColumnBuilderCategoryNode(string name, Func<ColumnBuilderPropertyNode, bool> filterPredicate)
        {
            Name = string.IsNullOrWhiteSpace(name) ? "Uncategorized" : name;
            Properties = new ObservableCollection<ColumnBuilderPropertyNode>();
            _propertiesView = CollectionViewSource.GetDefaultView(Properties);
            _filterPredicate = filterPredicate;
            _propertiesView.Filter = o =>
            {
                var prop = o as ColumnBuilderPropertyNode;
                if (prop == null) return false;
                return _filterPredicate?.Invoke(prop) ?? true;
            };
            _isExpanded = false;
        }

        public string Name { get; }

        public ObservableCollection<ColumnBuilderPropertyNode> Properties { get; }

        public ICollectionView PropertiesView => _propertiesView;

        public bool IsExpanded
        {
            get => _isExpanded;
            set => SetField(ref _isExpanded, value);
        }

        public bool IsVisible
        {
            get => _isVisible;
            set => SetField(ref _isVisible, value);
        }

        public bool? IsChecked
        {
            get => _isChecked;
            set
            {
                if (SetField(ref _isChecked, value))
                {
                    if (_suppressChildren) return;
                    if (value == null)
                    {
                        SelectionChanged?.Invoke();
                        return;
                    }

                    _suppressChildren = true;
                    foreach (var child in Properties)
                    {
                        child.SetCheckedFromParent(value == true);
                    }
                    _suppressChildren = false;
                    SelectionChanged?.Invoke();
                }
            }
        }

        public event Action SelectionChanged;
        public event PropertyChangedEventHandler PropertyChanged;

        public void RefreshPropertiesView() => _propertiesView?.Refresh();

        public void RefreshVisibility()
        {
            _propertiesView?.Refresh();
            IsVisible = _propertiesView == null || !_propertiesView.IsEmpty;
        }

        public void RefreshFromChildren()
        {
            if (_suppressChildren) return;

            var allTrue = Properties.All(p => p.IsChecked == true);
            var allFalse = Properties.All(p => p.IsChecked == false);
            var newValue = allTrue ? (bool?)true : allFalse ? (bool?)false : null;

            _suppressChildren = true;
            SetField(ref _isChecked, newValue, nameof(IsChecked));
            _suppressChildren = false;
        }

        public void ApplyUserToggle()
        {
            var target = _isChecked != true;
            SetCheckedFromUser(target);
        }

        private void SetCheckedFromUser(bool isChecked)
        {
            if (_suppressChildren) return;

            _suppressChildren = true;
            SetField(ref _isChecked, isChecked, nameof(IsChecked));
            _suppressChildren = false;

            foreach (var child in Properties)
            {
                child.SetCheckedFromParent(isChecked);
            }

            SelectionChanged?.Invoke();
        }

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string name = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return false;
            }

            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            return true;
        }
    }

    public sealed class ColumnBuilderPropertyNode : INotifyPropertyChanged
    {
        private bool? _isChecked;
        private bool _isRequired;
        private bool _isVisible = true;
        private int _orderIndex;

        public ColumnBuilderPropertyNode(string id, string propertyName, string categoryName)
        {
            Id = id;
            PropertyName = propertyName;
            CategoryName = categoryName;
        }

        public string Id { get; }
        public string PropertyName { get; }
        public string CategoryName { get; }
        public int SortIndex { get; set; }
        public int OrderIndex
        {
            get => _orderIndex;
            set => SetField(ref _orderIndex, value);
        }
        public ColumnBuilderCategoryNode Parent { get; set; }

        public bool IsVisible
        {
            get => _isVisible;
            set => SetField(ref _isVisible, value);
        }

        public bool? IsChecked
        {
            get => _isChecked;
            set
            {
                var normalized = value == true;
                if (SetField(ref _isChecked, normalized))
                {
                    if (!normalized && _isRequired)
                    {
                        SetField(ref _isRequired, false, nameof(IsRequired));
                    }
                    Parent?.RefreshFromChildren();
                    SelectionChanged?.Invoke();
                }
            }
        }

        public bool IsRequired
        {
            get => _isRequired;
            set => SetField(ref _isRequired, value);
        }

        public event Action SelectionChanged;
        public event PropertyChangedEventHandler PropertyChanged;

        public void SetCheckedFromParent(bool isChecked)
        {
            if (SetField(ref _isChecked, isChecked, nameof(IsChecked)))
            {
                SelectionChanged?.Invoke();
            }
        }

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string name = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return false;
            }

            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            return true;
        }
    }
}
