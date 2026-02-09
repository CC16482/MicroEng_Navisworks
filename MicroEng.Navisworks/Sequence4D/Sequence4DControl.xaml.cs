using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Autodesk.Navisworks.Api;
using WpfUiControls = Wpf.Ui.Controls;

namespace MicroEng.Navisworks
{
    public partial class Sequence4DControl : UserControl
    {
        private sealed class OrderingOption
        {
            public OrderingOption(Sequence4DOrdering value, string displayName)
            {
                Value = value;
                DisplayName = displayName;
            }

            public Sequence4DOrdering Value { get; }
            public string DisplayName { get; }
        }

        private static readonly List<OrderingOption> OrderingOptions = new()
        {
            new OrderingOption(Sequence4DOrdering.SelectionOrder, "Selection order"),
            new OrderingOption(Sequence4DOrdering.DistanceToReference, "Distance to reference"),
            new OrderingOption(Sequence4DOrdering.WorldXAscending, "World X (ascending)"),
            new OrderingOption(Sequence4DOrdering.WorldYAscending, "World Y (ascending)"),
            new OrderingOption(Sequence4DOrdering.WorldZAscending, "World Z (ascending)"),
            new OrderingOption(Sequence4DOrdering.PropertyValue, "Property value"),
            new OrderingOption(Sequence4DOrdering.Random, "Random")
        };

        private ModelItemCollection _captured = new ModelItemCollection();
        private ModelItem _reference;

        static Sequence4DControl()
        {
            AssemblyResolver.EnsureRegistered();
        }

        public Sequence4DControl()
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            Loaded += (_, __) => InitializeUi();
        }

        private void InitializeUi()
        {
            OrderingCombo.ItemsSource = OrderingOptions;
            OrderingCombo.DisplayMemberPath = nameof(OrderingOption.DisplayName);
            OrderingCombo.SelectedValuePath = nameof(OrderingOption.Value);
            OrderingCombo.SelectedValue = Sequence4DOrdering.DistanceToReference;

            TaskTypeCombo.ItemsSource = new[] { "Construct", "Demolish", "Temporary" };
            TaskTypeCombo.SelectedIndex = 0;

            var defaultStart = DateTime.Today.AddHours(8);
            StartDatePicker.SelectedDate = defaultStart.Date;
            StartTimeTextBox.Text = defaultStart.ToString("HH:mm:ss", CultureInfo.InvariantCulture);

            UpdateSelectionLabels();
            UpdateOrderingDependentFields();
            Log("Ready. Capture a selection to begin.");
        }

        private void OrderingCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateOrderingDependentFields();
        }

        private void CaptureSelection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                if (doc == null)
                {
                    throw new InvalidOperationException("No active document.");
                }

                if (doc.CurrentSelection == null || doc.CurrentSelection.IsEmpty)
                {
                    throw new InvalidOperationException("Selection is empty.");
                }

                var collection = new ModelItemCollection();
                doc.CurrentSelection.SelectedItems.CopyTo(collection);
                _captured = collection;

                UpdateSelectionLabels();
                Log($"Captured {_captured.Count} item(s).");
            }
            catch (Exception ex)
            {
                Log("Capture failed: " + ex.Message);
            }
        }

        private void SetReference_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                if (doc == null)
                {
                    throw new InvalidOperationException("No active document.");
                }

                if (doc.CurrentSelection == null || doc.CurrentSelection.IsEmpty)
                {
                    throw new InvalidOperationException("Select exactly one item to use as the reference.");
                }

                if (doc.CurrentSelection.SelectedItems.Count != 1)
                {
                    throw new InvalidOperationException("Select exactly one item to use as the reference.");
                }

                var reference = doc.CurrentSelection.SelectedItems.Cast<ModelItem>().FirstOrDefault();
                if (reference == null)
                {
                    throw new InvalidOperationException("Could not read selected item as reference.");
                }

                _reference = reference;
                UpdateSelectionLabels();
                Log($"Reference set: {_reference.DisplayName}");
            }
            catch (Exception ex)
            {
                Log("Set Reference failed: " + ex.Message);
            }
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var options = BuildOptions();
                var created = Sequence4DGenerator.GenerateTimelinerSequence(options);
                Log($"Done. Created {created} task(s) under root task \"{options.SequenceName}\".");
                Log("Open TimeLiner > Tasks/Simulate to play the sequence.");
                ShowSnackbar("Sequence generated",
                    $"Created {created} task(s).",
                    WpfUiControls.ControlAppearance.Success,
                    WpfUiControls.SymbolRegular.CheckmarkCircle24);
                FlashSuccess(sender as System.Windows.Controls.Button);
            }
            catch (Exception ex)
            {
                Log("Generate failed: " + ex.Message);
                ShowSnackbar("Generate failed",
                    ex.Message,
                    WpfUiControls.ControlAppearance.Danger,
                    WpfUiControls.SymbolRegular.ErrorCircle24);
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var name = SequenceNameTextBox.Text?.Trim();
                var removed = Sequence4DGenerator.DeleteTimelinerSequence(name);
                Log($"Deleted {removed} root sequence task(s) named \"{name}\".");
                ShowSnackbar("Sequence deleted",
                    $"Deleted {removed} task(s).",
                    WpfUiControls.ControlAppearance.Success,
                    WpfUiControls.SymbolRegular.CheckmarkCircle24);
                FlashSuccess(sender as System.Windows.Controls.Button);
            }
            catch (Exception ex)
            {
                Log("Delete failed: " + ex.Message);
                ShowSnackbar("Delete failed",
                    ex.Message,
                    WpfUiControls.ControlAppearance.Danger,
                    WpfUiControls.SymbolRegular.ErrorCircle24);
            }
        }

        private Sequence4DOptions BuildOptions()
        {
            return new Sequence4DOptions
            {
                SourceItems = _captured,
                ReferenceItem = _reference,
                Ordering = GetSelectedOrdering(),
                PropertyPath = PropertyPathTextBox.Text?.Trim(),
                SequenceName = SequenceNameTextBox.Text?.Trim(),
                TaskNamePrefix = TaskPrefixTextBox.Text?.Trim(),
                ItemsPerTask = ReadInt(ItemsPerTaskBox, 1),
                DurationSeconds = ReadDouble(DurationBox, 10.0, 0.1),
                OverlapSeconds = ReadDouble(OverlapBox, 0.0, 0.0),
                StartDateTime = GetStartDateTime(),
                SimulationTaskTypeName = TaskTypeCombo.SelectedItem?.ToString() ?? "Construct"
            };
        }

        private void ShowSnackbar(string title, string message, WpfUiControls.ControlAppearance appearance, WpfUiControls.SymbolRegular icon)
        {
            MicroEngSnackbar.Show(SnackbarPresenter, title, message, appearance, icon);
        }

        private void FlashSuccess(System.Windows.Controls.Button button)
        {
            if (button == null)
            {
                return;
            }

            var flashBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(76, 175, 80));
            var animation = new ColorAnimation
            {
                From = flashBrush.Color,
                To = System.Windows.Media.Colors.White,
                Duration = TimeSpan.FromMilliseconds(6000),
                BeginTime = TimeSpan.FromSeconds(1),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            animation.Completed += (_, _) => button.ClearValue(BackgroundProperty);
            button.Background = flashBrush;
            flashBrush.BeginAnimation(SolidColorBrush.ColorProperty, animation);
        }

        private Sequence4DOrdering GetSelectedOrdering()
        {
            if (OrderingCombo.SelectedValue is Sequence4DOrdering value)
            {
                return value;
            }

            if (OrderingCombo.SelectedItem is OrderingOption option)
            {
                return option.Value;
            }

            return Sequence4DOrdering.DistanceToReference;
        }

        private int ReadInt(Wpf.Ui.Controls.NumberBox box, int fallback)
        {
            if (box?.Value == null)
            {
                return fallback;
            }

            var value = (int)Math.Round(box.Value.Value, 0, MidpointRounding.AwayFromZero);
            return Math.Max(1, value);
        }

        private double ReadDouble(Wpf.Ui.Controls.NumberBox box, double fallback, double min)
        {
            if (box?.Value == null)
            {
                return fallback;
            }

            return Math.Max(min, box.Value.Value);
        }

        private DateTime GetStartDateTime()
        {
            var date = StartDatePicker.SelectedDate ?? DateTime.Today;
            var timeText = StartTimeTextBox.Text?.Trim();
            if (string.IsNullOrEmpty(timeText))
            {
                return date.Date;
            }

            if (TimeSpan.TryParse(timeText, CultureInfo.CurrentCulture, out var time)
                || TimeSpan.TryParse(timeText, CultureInfo.InvariantCulture, out time))
            {
                return date.Date.Add(time);
            }

            Log("Invalid time format. Use HH:mm:ss.");
            return date.Date.AddHours(8);
        }

        private void UpdateSelectionLabels()
        {
            var count = _captured?.Count ?? 0;
            SelectionStatusText.Text = $"Captured items: {count}";
            ReferenceStatusText.Text = _reference == null ? "Reference: (none)" : $"Reference: {_reference.DisplayName}";
        }

        private void UpdateOrderingDependentFields()
        {
            var ordering = GetSelectedOrdering();
            PropertyPathTextBox.IsEnabled = ordering == Sequence4DOrdering.PropertyValue;
            SetReferenceButton.IsEnabled = ordering == Sequence4DOrdering.DistanceToReference;
        }

        private void Log(string message)
        {
            if (LogTextBox == null)
            {
                return;
            }

            void Append()
            {
                LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
                LogTextBox.ScrollToEnd();
            }

            if (Dispatcher.CheckAccess())
            {
                Append();
            }
            else
            {
                Dispatcher.Invoke((Action)Append);
            }
        }
    }
}

