using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;

namespace MicroEng.Navisworks
{
    public partial class MicroEngPanelControl : UserControl
    {
        public MicroEngPanelControl()
        {
            InitializeComponent();

            if (!DesignerProperties.GetIsInDesignMode(this))
            {
                MicroEngActions.LogMessage += LogToPanel;
            }
        }

        private void AppendData_Click(object sender, RoutedEventArgs e)
        {
            LogToPanel("Opening Data Mapper...");
            MicroEngActions.AppendData();
            LogToPanel("Data Mapper closed.");
        }

        private void Reconstruct_Click(object sender, RoutedEventArgs e)
        {
            MicroEngActions.Reconstruct();
            LogToPanel("[Reconstruct] executed.");
        }

        private void ZoneFinder_Click(object sender, RoutedEventArgs e)
        {
            MicroEngActions.ZoneFinder();
            LogToPanel("[Zone Finder] executed.");
        }

        private void DataScraper_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var win = new DataScraperWindow();
                System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(win);
                win.Show();
            }
            catch (Exception ex)
            {
                LogToPanel($"[Data Scraper] failed to open: {ex.Message}");
                System.Windows.MessageBox.Show($"Data Scraper failed to open: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DataMatrix_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MicroEngActions.DataMatrix();
            }
            catch (Exception ex)
            {
                LogToPanel($"[Data Matrix] failed to open: {ex.Message}");
                System.Windows.MessageBox.Show($"Data Matrix failed to open: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LogToPanel(string message)
        {
            if (Dispatcher.CheckAccess())
            {
                LogTextBox.AppendText($"{DateTime.Now:HH:mm:ss} {message}{Environment.NewLine}");
                LogTextBox.ScrollToEnd();
            }
            else
            {
                Dispatcher.Invoke(() => LogToPanel(message));
            }
        }

        ~MicroEngPanelControl()
        {
            if (!DesignerProperties.GetIsInDesignMode(this))
            {
                MicroEngActions.LogMessage -= LogToPanel;
            }
        }
    }
}
