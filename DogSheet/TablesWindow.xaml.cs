using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace DogSheet
{
    /// <summary>
    /// Логика взаимодействия для TablesWindow.xaml
    /// </summary>
    public partial class TablesWindow : Window
    {
        private readonly MainWindow MW;

        public TablesWindow(MainWindow mainWindow)
        {
            MW = mainWindow;
            InitializeComponent();
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
