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
using Excel = Microsoft.Office.Interop.Excel;


namespace DogSheet
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Excel.Application MWapp;
        Excel.Workbook MWworkbook;
        public MainWindow()
        {
            MWapp = new Excel.Application();
            MWworkbook = MWapp.Workbooks.Open(@"C:\Users\pshar\source\repos\DogSheet\Журнал отлова безнадзорных животных.xlsx");
            InitializeComponent();
        }

        private void TablesButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"C:\Users\pshar\source\repos\DogSheet\Журнал отлова безнадзорных животных.xlsx");
        }

        private void DocsButton_Click(object sender, RoutedEventArgs e)
        {
            DocsWindow DW = new DocsWindow(MWapp, MWworkbook);
            DW.Show();
        }
    }
}
