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
            try
            {
                _ = System.Diagnostics.Process.Start(MW.pathToShort);
            }
            catch
            {
                MW.ShortChoice();
            }
        }

        private void FullTableButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _ = System.Diagnostics.Process.Start(MW.pathToFull);
            }
            catch
            {
                MW.FullChoice();
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MW.Show();
        }

        private void ChangeShortButton_Click(object sender, RoutedEventArgs e)
        {
            MW.ShortChoice();
        }

        private void ChangeFullButton_Click(object sender, RoutedEventArgs e)
        {
            MW.FullChoice();
        }
    }
}
