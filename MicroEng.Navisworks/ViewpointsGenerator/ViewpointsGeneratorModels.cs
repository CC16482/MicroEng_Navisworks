using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace MicroEng.Navisworks.ViewpointsGenerator
{
    public enum ViewpointsSourceMode
    {
        CurrentSelection = 0,
        SelectionSets = 1,
        PropertyGroups = 2
    }

    public enum ViewDirectionPreset
    {
        IsoSE = 0,
        IsoSW = 1,
        Top = 2,
        Bottom = 3,
        Front = 4,
        Back = 5,
        Left = 6,
        Right = 7
    }

    public enum ProjectionMode
    {
        Perspective = 0,
        Orthographic = 1
    }

    public sealed class ViewpointsGeneratorSettings : INotifyPropertyChanged
    {
        private ViewpointsSourceMode _sourceMode = ViewpointsSourceMode.SelectionSets;
        private string _outputFolderPath = "MicroEng/Viewpoints Generator";
        private string _namePrefix = "";
        private string _nameSuffix = "";
        private ViewDirectionPreset _direction = ViewDirectionPreset.IsoSE;
        private ProjectionMode _projection = ProjectionMode.Perspective;
        private double _fitMarginFactor = 0.15;

        public ViewpointsSourceMode SourceMode
        {
            get => _sourceMode;
            set
            {
                if (_sourceMode == value)
                {
                    return;
                }

                _sourceMode = value;
                OnPropertyChanged();
            }
        }

        public string OutputFolderPath
        {
            get => _outputFolderPath;
            set
            {
                if (string.Equals(_outputFolderPath, value, StringComparison.Ordinal))
                {
                    return;
                }

                _outputFolderPath = value ?? "";
                OnPropertyChanged();
            }
        }

        public string NamePrefix
        {
            get => _namePrefix;
            set
            {
                if (string.Equals(_namePrefix, value, StringComparison.Ordinal))
                {
                    return;
                }

                _namePrefix = value ?? "";
                OnPropertyChanged();
            }
        }

        public string NameSuffix
        {
            get => _nameSuffix;
            set
            {
                if (string.Equals(_nameSuffix, value, StringComparison.Ordinal))
                {
                    return;
                }

                _nameSuffix = value ?? "";
                OnPropertyChanged();
            }
        }

        public ViewDirectionPreset Direction
        {
            get => _direction;
            set
            {
                if (_direction == value)
                {
                    return;
                }

                _direction = value;
                OnPropertyChanged();
            }
        }

        public ProjectionMode Projection
        {
            get => _projection;
            set
            {
                if (_projection == value)
                {
                    return;
                }

                _projection = value;
                OnPropertyChanged();
            }
        }

        public double FitMarginFactor
        {
            get => _fitMarginFactor;
            set
            {
                var clamped = Math.Max(0, value);
                if (Math.Abs(_fitMarginFactor - clamped) < 0.000001)
                {
                    return;
                }

                _fitMarginFactor = clamped;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string prop = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }

    public sealed class ViewpointPlanItem : INotifyPropertyChanged
    {
        private bool _enabled = true;
        private string _name = "";
        private string _source = "";
        private int _itemCount;

        public bool Enabled
        {
            get => _enabled;
            set
            {
                if (_enabled == value)
                {
                    return;
                }

                _enabled = value;
                OnPropertyChanged();
            }
        }

        public string Name
        {
            get => _name;
            set
            {
                if (string.Equals(_name, value, StringComparison.Ordinal))
                {
                    return;
                }

                _name = value ?? "";
                OnPropertyChanged();
            }
        }

        public string Source
        {
            get => _source;
            set
            {
                if (string.Equals(_source, value, StringComparison.Ordinal))
                {
                    return;
                }

                _source = value ?? "";
                OnPropertyChanged();
            }
        }

        public int ItemCount
        {
            get => _itemCount;
            set
            {
                if (_itemCount == value)
                {
                    return;
                }

                _itemCount = value;
                OnPropertyChanged();
            }
        }

        internal Func<Autodesk.Navisworks.Api.ModelItemCollection> ResolveItems { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string prop = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }

    public sealed class SelectionSetPickerItem : INotifyPropertyChanged
    {
        private bool _enabled = true;
        private string _path = "";
        private Autodesk.Navisworks.Api.SelectionSet _set;

        public bool Enabled
        {
            get => _enabled;
            set
            {
                if (_enabled == value)
                {
                    return;
                }

                _enabled = value;
                OnPropertyChanged();
            }
        }

        public string Path
        {
            get => _path;
            set
            {
                if (string.Equals(_path, value, StringComparison.Ordinal))
                {
                    return;
                }

                _path = value ?? "";
                OnPropertyChanged();
            }
        }

        public Autodesk.Navisworks.Api.SelectionSet Set
        {
            get => _set;
            set
            {
                if (ReferenceEquals(_set, value))
                {
                    return;
                }

                _set = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string prop = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
