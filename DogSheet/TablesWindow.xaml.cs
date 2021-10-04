using System;
using System.Windows;

namespace DogSheet
{
    public partial class TablesWindow : Window
    {
        private readonly MainWindow MW;

        public TablesWindow(MainWindow mainWindow)
        {
            InitializeComponent();

            MW = mainWindow;
        }

        private void ShortTableButton_Click(object sender, RoutedEventArgs e)
        {
            _ = System.Diagnostics.Process.Start(MW.pathToShort);
        }

        private void FullTableButton_Click(object sender, RoutedEventArgs e)
        {
            _ = System.Diagnostics.Process.Start(MW.pathToFull);
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MW.Show();
        }
    }
}
