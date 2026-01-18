using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.SmartSets
{
    public partial class PropertyPickerWindow : Window
    {
        public PropertyPickerViewModel VM { get; }

        public ScrapedPropertyDescriptor Selected => VM.SelectedProperty;

        public PropertyPickerWindow(PropertyPickerViewModel vm)
        {
            InitializeComponent();
            VM = vm ?? throw new ArgumentNullException(nameof(vm));
            DataContext = VM;

            try
            {
                MicroEngWpfUiTheme.ApplyTo(this);
            }
            catch
            {
                // ignore theme failures
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (VM.SelectedProperty == null)
            {
                return;
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void DataGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (VM.SelectedProperty == null)
            {
                return;
            }

            DialogResult = true;
            Close();
        }
    }

    public sealed class PropertyPickerViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<string> Categories { get; } = new ObservableCollection<string>();
        public ObservableCollection<ScrapedPropertyDescriptor> AllProperties { get; } = new ObservableCollection<ScrapedPropertyDescriptor>();
        public ObservableCollection<ScrapedPropertyDescriptor> FilteredProperties { get; } = new ObservableCollection<ScrapedPropertyDescriptor>();

        private string _searchText = "";
        public string SearchText
        {
            get => _searchText;
            set
            {
                _searchText = value ?? "";
                OnPropertyChanged();
                Refilter();
            }
        }

        private string _selectedCategory;
        public string SelectedCategory
        {
            get => _selectedCategory;
            set
            {
                _selectedCategory = value;
                OnPropertyChanged();
                Refilter();
            }
        }

        private ScrapedPropertyDescriptor _selectedProperty;
        public ScrapedPropertyDescriptor SelectedProperty
        {
            get => _selectedProperty;
            set
            {
                _selectedProperty = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanAccept));
            }
        }

        public bool CanAccept => SelectedProperty != null;

        public PropertyPickerViewModel(System.Collections.Generic.IEnumerable<ScrapedPropertyDescriptor> props)
        {
            var list = (props ?? Enumerable.Empty<ScrapedPropertyDescriptor>()).ToList();

            foreach (var p in list)
            {
                AllProperties.Add(p);
            }

            foreach (var c in list.Select(p => p.Category).Distinct().OrderBy(s => s))
            {
                Categories.Add(c);
            }

            SelectedCategory = Categories.FirstOrDefault();
            Refilter();
        }

        private void Refilter()
        {
            var query = (SearchText ?? "").Trim();
            var catFilter = SelectedCategory;

            var items = AllProperties.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(catFilter))
            {
                items = items.Where(p => string.Equals(p.Category, catFilter, StringComparison.OrdinalIgnoreCase));
            }

            if (!string.IsNullOrWhiteSpace(query))
            {
                items = items.Where(p =>
                    (p.Category ?? "").IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (p.Name ?? "").IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            var result = items.OrderBy(p => p.Category).ThenBy(p => p.Name).ToList();

            FilteredProperties.Clear();
            foreach (var item in result)
            {
                FilteredProperties.Add(item);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
