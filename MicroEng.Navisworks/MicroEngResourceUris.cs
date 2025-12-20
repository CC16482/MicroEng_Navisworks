using System;

namespace MicroEng.Navisworks
{
    /// <summary>
    /// Provides resource URIs for XAML ResourceDictionary merges.
    ///
    /// IMPORTANT:
    /// Use pack URIs for merged dictionaries (BAML resources). File-path URIs are not reliable in
    /// Navisworks hosting and can cause StaticResource lookups to fail at runtime.
    /// </summary>
    public static class MicroEngResourceUris
    {
        public static readonly Uri WpfUiRoot = new Uri(
            "pack://application:,,,/MicroEng.Navisworks;component/Assets/Theme/MicroEngWpfUiRoot.xaml",
            UriKind.Absolute);
    }
}
