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
using Microsoft.Win32;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;


namespace DogSheet
{
    public partial class MainWindow : Window
    {
        public Excel.Application exApp;
        public Excel.Workbook exWorkbook;
        public string pathToShort = @"C:\Users\pshar\source\repos\DogSheet\Журнал отлова (short).xlsx";
        public string pathToFull = @"C:\Users\pshar\source\repos\DogSheet\Журнал отлова (full).xlsx";

        public MainWindow()                                                         //Главное окно
        {
            exApp = new Excel.Application();
            exWorkbook = exApp.Workbooks.Open(pathToShort);
            InitializeComponent();
        }

        private void TablesButton_Click(object sender, RoutedEventArgs e)           //Открытие таблиц для просмотра (заменить на выбор)
        {
            TablesWindow TW = new TablesWindow(this);
            TW.Show();
            this.Hide();
        }

        private void DocsButton_Click(object sender, RoutedEventArgs e)             //Открытие окна для создания отчетов
        {
            DocsWindow DW = new DocsWindow(this);
            DW.Show();
            this.Hide();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            exWorkbook.Close(true);
            exApp.Quit();
        }
    }
}
