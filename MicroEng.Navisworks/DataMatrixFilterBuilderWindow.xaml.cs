using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Collections.ObjectModel;
using System.Collections;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace MicroEng.Navisworks
{
    internal sealed partial class DataMatrixFilterBuilderWindow
    {
        private readonly ObservableCollection<FilterRuleViewModel> _rules;

        public string ColumnId { get; }
        public string ColumnLabel { get; }
        public ObservableCollection<OperatorOption> OperatorOptions { get; }
        public ObservableCollection<JoinOption> JoinOptions { get; }

        public List<DataMatrixColumnFilter> Filters { get; private set; } = new List<DataMatrixColumnFilter>();
        private readonly bool _caseSensitive;
        private readonly bool _trimWhitespace;

        public bool FilterCaseSensitive => _caseSensitive;
        public bool FilterTrimWhitespace => _trimWhitespace;

        public DataMatrixFilterBuilderWindow(
            IEnumerable<DataMatrixColumnFilter> existingFilters,
            bool caseSensitive,
            bool trimWhitespace,
            string columnId,
            string columnLabel)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            SetResourceReference(BackgroundProperty, "ApplicationBackgroundBrush");
            SetResourceReference(ForegroundProperty, "TextFillColorPrimaryBrush");

            ColumnId = columnId ?? string.Empty;
            var label = string.IsNullOrWhiteSpace(columnLabel) ? ColumnId : columnLabel;
            ColumnLabel = string.IsNullOrWhiteSpace(label) ? "Filtering" : $"Filtering: {label}";
            if (!string.IsNullOrWhiteSpace(label))
            {
                Title = $"Filter Builder - {label}";
            }
            OperatorOptions = new ObservableCollection<OperatorOption>(BuildOperatorOptions());
            JoinOptions = new ObservableCollection<JoinOption>(BuildJoinOptions());
            DataContext = this;

            _caseSensitive = caseSensitive;
            _trimWhitespace = trimWhitespace;

            _rules = new ObservableCollection<FilterRuleViewModel>(
                (existingFilters ?? Enumerable.Empty<DataMatrixColumnFilter>())
                .Where(f => f != null)
                .Select(FilterRuleViewModel.FromFilter));
            RulesGrid.ItemsSource = _rules;
            HookRuleEvents();

            AddRuleButton.Click += AddRuleButton_Click;
            ApplyButton.Click += ApplyButton_Click;
            CancelButton.Click += (_, __) => DialogResult = false;
        }

        private void HookRuleEvents()
        {
            if (_rules == null) return;
            _rules.CollectionChanged += Rules_CollectionChanged;
            foreach (var rule in _rules)
            {
                rule.PropertyChanged += Rule_PropertyChanged;
            }
            RefreshJoinAvailability();
        }

        private void Rules_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (var item in e.OldItems.OfType<FilterRuleViewModel>())
                {
                    item.PropertyChanged -= Rule_PropertyChanged;
                }
            }

            if (e.NewItems != null)
            {
                foreach (var item in e.NewItems.OfType<FilterRuleViewModel>())
                {
                    item.PropertyChanged += Rule_PropertyChanged;
                }
            }

            RefreshJoinAvailability();
        }

        private void Rule_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(FilterRuleViewModel.IsEnabled))
            {
                RefreshJoinAvailability();
            }
        }

        private void RefreshJoinAvailability()
        {
            if (RulesGrid?.Items != null)
            {
                RulesGrid.Items.Refresh();
            }
        }

        private static IEnumerable<OperatorOption> BuildOperatorOptions()
        {
            return new[]
            {
                new OperatorOption(DataMatrixFilterOperator.Contains, "Contains"),
                new OperatorOption(DataMatrixFilterOperator.NotContains, "Does not contain"),
                new OperatorOption(DataMatrixFilterOperator.Equals, "Equals"),
                new OperatorOption(DataMatrixFilterOperator.NotEquals, "Not equals"),
                new OperatorOption(DataMatrixFilterOperator.StartsWith, "Starts with"),
                new OperatorOption(DataMatrixFilterOperator.EndsWith, "Ends with"),
                new OperatorOption(DataMatrixFilterOperator.Wildcard, "Matches wildcard"),
                new OperatorOption(DataMatrixFilterOperator.Regex, "Matches regex"),
                new OperatorOption(DataMatrixFilterOperator.GreaterThan, "Greater than"),
                new OperatorOption(DataMatrixFilterOperator.GreaterOrEqual, "Greater or equal"),
                new OperatorOption(DataMatrixFilterOperator.LessThan, "Less than"),
                new OperatorOption(DataMatrixFilterOperator.LessOrEqual, "Less or equal"),
                new OperatorOption(DataMatrixFilterOperator.Between, "Between"),
                new OperatorOption(DataMatrixFilterOperator.DateBefore, "Date before"),
                new OperatorOption(DataMatrixFilterOperator.DateAfter, "Date after"),
                new OperatorOption(DataMatrixFilterOperator.DateOn, "Date on"),
                new OperatorOption(DataMatrixFilterOperator.DateBetween, "Date between"),
                new OperatorOption(DataMatrixFilterOperator.Blank, "Is blank"),
                new OperatorOption(DataMatrixFilterOperator.NotBlank, "Is not blank"),
                new OperatorOption(DataMatrixFilterOperator.IsEmpty, "Is empty"),
                new OperatorOption(DataMatrixFilterOperator.IsNotEmpty, "Is not empty"),
                new OperatorOption(DataMatrixFilterOperator.IsNull, "Is null"),
                new OperatorOption(DataMatrixFilterOperator.IsNotNull, "Is not null"),
                new OperatorOption(DataMatrixFilterOperator.InList, "In list")
            };
        }

        private static IEnumerable<JoinOption> BuildJoinOptions()
        {
            return new[]
            {
                new JoinOption(DataMatrixFilterJoin.Or, "OR"),
                new JoinOption(DataMatrixFilterJoin.And, "AND")
            };
        }

        private void AddRuleButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorText.Text = string.Empty;
            var op = DataMatrixFilterOperator.Contains;
            var value = string.Empty;

            if (string.IsNullOrWhiteSpace(ColumnId))
            {
                ErrorText.Text = "Column is unavailable for this rule.";
                return;
            }

            var rule = new FilterRuleViewModel
            {
                AttributeId = ColumnId,
                Operator = op,
                Value = value,
                IsEnabled = true,
                CaseSensitive = _caseSensitive,
                TrimWhitespace = _trimWhitespace,
                Join = DataMatrixFilterJoin.Or
            };

            _rules.Add(rule);
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorText.Text = string.Empty;
            var error = ValidateRules();
            if (!string.IsNullOrWhiteSpace(error))
            {
                ErrorText.Text = error;
                return;
            }

            Filters = _rules.Select(r => r.ToFilter()).ToList();
            DialogResult = true;
        }

        private string ValidateRules()
        {
            foreach (var rule in _rules.Where(r => r.IsEnabled))
            {
                if (string.IsNullOrWhiteSpace(rule.AttributeId))
                {
                    return "Every rule needs a column.";
                }

                if (RequiresValue(rule.Operator) && string.IsNullOrWhiteSpace(rule.Value))
                {
                    if (IsNumericOperator(rule.Operator))
                    {
                        return $"Value required for \"{GetOperatorLabel(rule.Operator)}\".";
                    }

                    if (IsDateOperator(rule.Operator))
                    {
                        return $"Value required for \"{GetOperatorLabel(rule.Operator)}\".";
                    }
                }

                if (rule.Operator == DataMatrixFilterOperator.Regex && !string.IsNullOrWhiteSpace(rule.Value))
                {
                    if (!IsValidRegex(rule.Value, FilterCaseSensitive))
                    {
                        return $"Invalid regex: {rule.Value}";
                    }
                }

                if (IsNumericOperator(rule.Operator))
                {
                    if (!TryParseNumber(rule.Value, out _, out var rangeError))
                    {
                        return rangeError ?? $"Invalid number for \"{GetOperatorLabel(rule.Operator)}\".";
                    }
                }

                if (IsDateOperator(rule.Operator))
                {
                    if (!TryParseDate(rule.Value, out _, out var dateError))
                    {
                        return dateError ?? $"Invalid date for \"{GetOperatorLabel(rule.Operator)}\".";
                    }
                }
            }

            return null;
        }


        private static bool IsNumericOperator(DataMatrixFilterOperator op)
        {
            return op == DataMatrixFilterOperator.GreaterThan
                   || op == DataMatrixFilterOperator.GreaterOrEqual
                   || op == DataMatrixFilterOperator.LessThan
                   || op == DataMatrixFilterOperator.LessOrEqual
                   || op == DataMatrixFilterOperator.Between;
        }

        private static bool IsDateOperator(DataMatrixFilterOperator op)
        {
            return op == DataMatrixFilterOperator.DateBefore
                   || op == DataMatrixFilterOperator.DateAfter
                   || op == DataMatrixFilterOperator.DateOn
                   || op == DataMatrixFilterOperator.DateBetween;
        }

        private static bool RequiresValue(DataMatrixFilterOperator op)
        {
            return op != DataMatrixFilterOperator.IsEmpty
                   && op != DataMatrixFilterOperator.IsNotEmpty
                   && op != DataMatrixFilterOperator.IsNull
                   && op != DataMatrixFilterOperator.IsNotNull
                   && op != DataMatrixFilterOperator.Blank
                   && op != DataMatrixFilterOperator.NotBlank;
        }

        private static string GetOperatorLabel(DataMatrixFilterOperator op)
        {
            return op.ToString();
        }

        private static bool IsValidRegex(string pattern, bool caseSensitive)
        {
            try
            {
                var options = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
                _ = new Regex(pattern, options);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryParseNumber(string text, out double value, out string error)
        {
            error = null;
            value = 0d;

            if (string.IsNullOrWhiteSpace(text))
            {
                error = "Enter a number.";
                return false;
            }

            if (text.Contains(".."))
            {
                var parts = text.Split(new[] { ".." }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length != 2)
                {
                    error = "Use min..max for between.";
                    return false;
                }
                if (!double.TryParse(parts[0], NumberStyles.Any, CultureInfo.CurrentCulture, out _)
                    && !double.TryParse(parts[0], NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                {
                    error = "Invalid minimum value.";
                    return false;
                }
                if (!double.TryParse(parts[1], NumberStyles.Any, CultureInfo.CurrentCulture, out _)
                    && !double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                {
                    error = "Invalid maximum value.";
                    return false;
                }
                return true;
            }

            if (!double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value)
                && !double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            {
                error = "Invalid number.";
                return false;
            }

            return true;
        }

        private static bool TryParseDate(string text, out DateTime value, out string error)
        {
            error = null;
            value = default;

            if (string.IsNullOrWhiteSpace(text))
            {
                error = "Enter a date.";
                return false;
            }

            if (text.Contains(".."))
            {
                var parts = text.Split(new[] { ".." }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length != 2)
                {
                    error = "Use start..end for between.";
                    return false;
                }
                if (!DateTime.TryParse(parts[0], CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out _)
                    && !DateTime.TryParse(parts[0], CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out _))
                {
                    error = "Invalid start date.";
                    return false;
                }
                if (!DateTime.TryParse(parts[1], CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out _)
                    && !DateTime.TryParse(parts[1], CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out _))
                {
                    error = "Invalid end date.";
                    return false;
                }
                return true;
            }

            if (!DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out value)
                && !DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out value))
            {
                error = "Invalid date.";
                return false;
            }

            return true;
        }

        private void DeleteRuleButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is FilterRuleViewModel rule)
            {
                _rules.Remove(rule);
            }
        }
    }

    internal sealed class OperatorOption
    {
        public OperatorOption(DataMatrixFilterOperator op, string label)
        {
            Operator = op;
            Label = label;
        }

        public DataMatrixFilterOperator Operator { get; }
        public string Label { get; }
    }

    internal sealed class JoinOption
    {
        public JoinOption(DataMatrixFilterJoin join, string label)
        {
            Join = join;
            Label = label;
        }

        public DataMatrixFilterJoin Join { get; }
        public string Label { get; }
    }

    internal sealed class FilterRuleViewModel : INotifyPropertyChanged
    {
        private string _attributeId;
        private DataMatrixFilterOperator _operator = DataMatrixFilterOperator.Contains;
        private string _value;
        private bool _isEnabled = true;
        private bool _caseSensitive;
        private bool _trimWhitespace = true;
        private DataMatrixFilterJoin _join = DataMatrixFilterJoin.Or;

        public string AttributeId
        {
            get => _attributeId;
            set
            {
                if (_attributeId == value) return;
                _attributeId = value;
                OnPropertyChanged();
            }
        }

        public DataMatrixFilterOperator Operator
        {
            get => _operator;
            set
            {
                if (_operator == value) return;
                _operator = value;
                OnPropertyChanged();
            }
        }

        public string Value
        {
            get => _value;
            set
            {
                if (_value == value) return;
                _value = value;
                OnPropertyChanged();
            }
        }

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                if (_isEnabled == value) return;
                _isEnabled = value;
                OnPropertyChanged();
            }
        }

        public bool CaseSensitive
        {
            get => _caseSensitive;
            set
            {
                if (_caseSensitive == value) return;
                _caseSensitive = value;
                OnPropertyChanged();
            }
        }

        public bool TrimWhitespace
        {
            get => _trimWhitespace;
            set
            {
                if (_trimWhitespace == value) return;
                _trimWhitespace = value;
                OnPropertyChanged();
            }
        }

        public DataMatrixFilterJoin Join
        {
            get => _join;
            set
            {
                if (_join == value) return;
                _join = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public static FilterRuleViewModel FromFilter(DataMatrixColumnFilter filter)
        {
            return new FilterRuleViewModel
            {
                AttributeId = filter.AttributeId,
                Operator = filter.Operator,
                Value = filter.Value ?? string.Empty,
                IsEnabled = filter.IsEnabled,
                CaseSensitive = filter.CaseSensitive,
                TrimWhitespace = filter.TrimWhitespace,
                Join = filter.Join
            };
        }

        public DataMatrixColumnFilter ToFilter()
        {
            return new DataMatrixColumnFilter
            {
                AttributeId = AttributeId,
                Operator = Operator,
                Value = Value ?? string.Empty,
                IsEnabled = IsEnabled,
                CaseSensitive = CaseSensitive,
                TrimWhitespace = TrimWhitespace,
                Join = Join
            };
        }

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    internal sealed class FilterValueEnabledConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is DataMatrixFilterOperator op)
            {
                return op != DataMatrixFilterOperator.IsEmpty
                       && op != DataMatrixFilterOperator.IsNotEmpty
                       && op != DataMatrixFilterOperator.IsNull
                       && op != DataMatrixFilterOperator.IsNotNull
                       && op != DataMatrixFilterOperator.Blank
                       && op != DataMatrixFilterOperator.NotBlank;
            }

            return true;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    internal sealed class JoinEnabledConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 2)
            {
                return true;
            }

            if (values[1] is not int index || index <= 0)
            {
                return false;
            }

            var list = values[0] as IList;
            if (list != null)
            {
                for (var i = 0; i < index && i < list.Count; i++)
                {
                    if (list[i] is FilterRuleViewModel rule && rule.IsEnabled)
                    {
                        return true;
                    }
                }
                return false;
            }

            if (values[0] is IEnumerable enumerable)
            {
                var i = 0;
                foreach (var item in enumerable)
                {
                    if (i >= index) break;
                    if (item is FilterRuleViewModel rule && rule.IsEnabled)
                    {
                        return true;
                    }
                    i++;
                }
                return false;
            }

            return index > 0;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
