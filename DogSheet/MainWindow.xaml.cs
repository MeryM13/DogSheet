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
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace DogSheet
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public DocsWindow DW = new DocsWindow();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void TablesButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"C:\Users\pshar\source\repos\DogSheet\Журнал отлова безнадзорных животных.xlsx");
        }

        private void DocsButton_Click(object sender, RoutedEventArgs e)
        {
            DW.Show();
        }
    }
}
