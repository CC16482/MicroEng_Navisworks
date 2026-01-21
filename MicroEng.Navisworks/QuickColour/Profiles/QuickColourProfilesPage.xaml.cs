using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace MicroEng.Navisworks.QuickColour.Profiles
{
    public partial class QuickColourProfilesPage : UserControl, INotifyPropertyChanged
    {
        private readonly MicroEngColourProfileStore _store = MicroEngColourProfileStore.CreateDefault();

        public ObservableCollection<MicroEngColourProfile> Profiles { get; } =
            new ObservableCollection<MicroEngColourProfile>();

        private MicroEngColourProfile _selectedProfile;
        public MicroEngColourProfile SelectedProfile
        {
            get => _selectedProfile;
            set
            {
                if (!ReferenceEquals(_selectedProfile, value))
                {
                    _selectedProfile = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(HasSelection));
                }
            }
        }

        public bool HasSelection => SelectedProfile != null;

        private string _statusText = "Ready.";
        public string StatusText
        {
            get => _statusText;
            set { _statusText = value ?? ""; OnPropertyChanged(); }
        }

        public event Action<MicroEngColourProfile, MicroEngColourApplyMode> ApplyRequested;

        public QuickColourProfilesPage()
        {
            InitializeComponent();
            DataContext = this;
            Refresh();
        }

        public void Refresh()
        {
            Profiles.Clear();
            foreach (var profile in _store.LoadAll())
            {
                Profiles.Add(profile);
            }

            StatusText = $"Loaded {Profiles.Count} profile(s).";
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            Refresh();
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedProfile == null) return;

            var name = SelectedProfile.Name ?? "(Unnamed)";
            if (MessageBox.Show($"Delete profile '{name}'?", "Delete Profile", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            _store.Delete(SelectedProfile.Id);
            Refresh();
            StatusText = $"Deleted '{name}'.";
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedProfile == null) return;

            var sfd = new SaveFileDialog
            {
                Title = "Export MicroEng Colour Profile",
                Filter = "MicroEng Colour Profile (*.json)|*.json|All files (*.*)|*.*",
                FileName = $"{SelectedProfile.Name}.json"
            };

            if (sfd.ShowDialog() != true) return;

            _store.ExportTo(SelectedProfile, sfd.FileName);
            StatusText = $"Exported '{SelectedProfile.Name}'.";
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Title = "Import MicroEng Colour Profile",
                Filter = "MicroEng Colour Profile (*.json)|*.json|All files (*.*)|*.*",
                Multiselect = false
            };

            if (ofd.ShowDialog() != true) return;

            try
            {
                _store.ImportFrom(ofd.FileName);
                Refresh();
                StatusText = $"Imported '{Path.GetFileName(ofd.FileName)}'.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Import failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ApplyTemporary_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedProfile == null) return;
            ApplyRequested?.Invoke(SelectedProfile, MicroEngColourApplyMode.Temporary);
        }

        private void ApplyPermanent_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedProfile == null) return;
            ApplyRequested?.Invoke(SelectedProfile, MicroEngColourApplyMode.Permanent);
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
