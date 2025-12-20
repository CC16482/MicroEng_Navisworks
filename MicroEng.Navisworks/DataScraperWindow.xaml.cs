using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Autodesk.Navisworks.Api;

namespace MicroEng.Navisworks
{
    public partial class DataScraperWindow : Window
    {
        static DataScraperWindow()
        {
            AssemblyResolver.EnsureRegistered();
        }

        private readonly DataScraperService _service = new();
        private List<ScrapedPropertyView> _currentProperties = new();
        private ScrapeScopeType _lastScope = ScrapeScopeType.CurrentSelection;
        private string _lastSelectionSet;
        private string _lastSearchSet;
        private List<RawRow> _rawRows = new();

        public DataScraperWindow(string initialProfile = null)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);
            try
            {
                if (!string.IsNullOrWhiteSpace(initialProfile))
                {
                    ProfileNameBox.Text = initialProfile;
                }
                LoadSets();
                RefreshHistory();
                RefreshProfile();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"DataScraper init failed: {ex.Message}");
                StatusText.Text = ex.Message;
            }
        }

        private void LoadSets()
        {
            try
            {
                var doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                if (doc == null)
                {
                    SelectionSetCombo.ItemsSource = new List<string>();
                    SearchSetCombo.ItemsSource = new List<string>();
                    return;
                }
                // Selection/Search sets not yet implemented; leave empty for now
                SelectionSetCombo.ItemsSource = new List<string>();
                SearchSetCombo.ItemsSource = new List<string>();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"DataScraper: failed to load sets: {ex.Message}");
            }
        }

        private void RefreshHistory()
        {
            HistoryList.ItemsSource = DataScraperCache.AllSessions
                .OrderByDescending(s => s.Timestamp)
                .ToList();
        }

        private void RefreshProfile()
        {
            // Always show the latest run for each unique profile (no filter)
            var sessions = DataScraperCache.AllSessions.AsEnumerable();
            var latestPerProfile = sessions
                .GroupBy(s => string.IsNullOrWhiteSpace(s.ProfileName) ? "Default" : s.ProfileName, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderByDescending(s => s.Timestamp).First())
                .OrderBy(s => s.ProfileName)
                .ToList();

            if (latestPerProfile.Any())
            {
                ProfileList.ItemsSource = latestPerProfile;
                ProfileSummaryText.Text = $"Latest runs across {latestPerProfile.Count} profile(s).";
            }
            else
            {
                ProfileList.ItemsSource = null;
                ProfileSummaryText.Text = "No profile runs yet.";
            }
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var profile = string.IsNullOrWhiteSpace(ProfileNameBox.Text) ? "Default" : ProfileNameBox.Text.Trim();
                var scope = GetScopeType();
                var items = _service.ResolveScope(scope, _lastSelectionSet, _lastSearchSet, out var description).ToList();
                if (!items.Any())
                {
                    MessageBox.Show("No items found for the selected scope.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                StatusText.Text = "Scraping...";
                var session = _service.Scrape(profile, scope, description, items);
                StatusText.Text = $"Scraped {session.ItemsScanned} items, {session.Properties.Count} properties.";
                MicroEngActions.Log($"Data Scraper: {session.ItemsScanned} items scanned in {description}. {session.Properties.Count} properties cached for profile '{profile}'.");
                RefreshHistory();
                RefreshProfile();
                ShowProperties(session);
            }
            catch (Exception ex)
            {
                StatusText.Text = ex.Message;
                MessageBox.Show(ex.Message, "MicroEng", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private ScrapeScopeType GetScopeType()
        {
            if (SingleItemRadio.IsChecked == true) return ScrapeScopeType.SingleItem;
            if (CurrentSelectionRadio.IsChecked == true) return ScrapeScopeType.CurrentSelection;
            if (SelectionSetRadio.IsChecked == true) return ScrapeScopeType.SelectionSet;
            if (SearchSetRadio.IsChecked == true) return ScrapeScopeType.SearchSet;
            if (EntireModelRadio.IsChecked == true) return ScrapeScopeType.EntireModel;
            return ScrapeScopeType.CurrentSelection;
        }

        private void HistoryList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (HistoryList.SelectedItem is ScrapeSession session)
            {
                ShowProperties(session);
            }
        }

        private void ShowProperties(ScrapeSession session)
        {
            _currentProperties = session.Properties
                .Select(p => new ScrapedPropertyView(p))
                .ToList();
            PropertiesGrid.ItemsSource = _currentProperties;
            PropertySummary.Text = $"{_currentProperties.Count} properties from session {session.Timestamp:T}";

            _rawRows = session.RawEntries
                .Select(r => new RawRow
                {
                    Profile = r.Profile,
                    Scope = r.Scope,
                    ItemPath = r.ItemPath,
                    Category = r.Category,
                    Name = r.Name,
                    DataType = r.DataType,
                    Value = r.Value
                }).ToList();
            RawDataGrid.ItemsSource = _rawRows;
            RawSummaryText.Text = $"{_rawRows.Count} raw entries.";
        }

        private void FilterBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_currentProperties == null) return;
            var text = FilterBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(text))
            {
                PropertiesGrid.ItemsSource = _currentProperties;
            }
            else
            {
                PropertiesGrid.ItemsSource = _currentProperties
                    .Where(p => p.Name.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                p.Category.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
            }
        }

        private void RawFilterBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_rawRows == null) return;
            var text = RawFilterBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(text))
            {
                RawDataGrid.ItemsSource = _rawRows;
            }
            else
            {
                RawDataGrid.ItemsSource = _rawRows
                    .Where(r => (r.Category?.IndexOf(text, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0
                             || (r.Name?.IndexOf(text, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0
                             || (r.Value?.IndexOf(text, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0
                             || (r.ItemPath?.IndexOf(text, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0)
                    .ToList();
            }
            RawSummaryText.Text = $"{(RawDataGrid.ItemsSource as System.Collections.ICollection)?.Count} raw entries.";
        }

        private void Rerun_Click(object sender, RoutedEventArgs e)
        {
            if (HistoryList.SelectedItem is ScrapeSession session)
            {
                _lastScope = Enum.TryParse<ScrapeScopeType>(session.ScopeType, out var scope) ? scope : ScrapeScopeType.CurrentSelection;
                Run_Click(sender, e);
            }
        }

        private void ViewProperties_Click(object sender, RoutedEventArgs e)
        {
            if (HistoryList.SelectedItem is ScrapeSession session)
            {
                ShowProperties(session);
            }
        }

        private void RefreshSets_Click(object sender, RoutedEventArgs e)
        {
            LoadSets();
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void SelectionSetCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _lastSelectionSet = SelectionSetCombo.SelectedItem as string;
        }

        private void SearchSetCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _lastSearchSet = SearchSetCombo.SelectedItem as string;
        }

        private void ProfileNameBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Profile name text is used when running a scrape; it no longer filters the Profile tab.
        }

        private void SelectionSetRadio_Checked(object sender, RoutedEventArgs e)
        {
            SelectionSetCombo.IsEnabled = true;
            SearchSetCombo.IsEnabled = false;
        }

        private void SearchSetRadio_Checked(object sender, RoutedEventArgs e)
        {
            SelectionSetCombo.IsEnabled = false;
            SearchSetCombo.IsEnabled = true;
        }

        private void OtherScope_Checked(object sender, RoutedEventArgs e)
        {
            SelectionSetCombo.IsEnabled = false;
            SearchSetCombo.IsEnabled = false;
        }
    }

    internal class ScrapedPropertyView
    {
        public ScrapedPropertyView(ScrapedProperty p)
        {
            Category = p.Category;
            Name = p.Name;
            DataType = p.DataType;
            ItemCount = p.ItemCount;
            DistinctValueCount = p.DistinctValueCount;
            SampleValuesString = string.Join("; ", p.SampleValues ?? new List<string>());
        }
        public string Category { get; }
        public string Name { get; }
        public string DataType { get; }
        public int ItemCount { get; }
        public int DistinctValueCount { get; }
        public string SampleValuesString { get; }
    }

    internal class RawRow
    {
        public string Profile { get; set; }
        public string Scope { get; set; }
        public string ItemPath { get; set; }
        public string Category { get; set; }
        public string Name { get; set; }
        public string DataType { get; set; }
        public string Value { get; set; }
    }
}
