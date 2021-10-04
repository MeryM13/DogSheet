using System.IO;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;


namespace DogSheet
{
    public partial class MainWindow : Window
    {
        public Excel.Application exApp;
        public Excel.Workbook exWorkbook;

        public string path;
        public string pathToShort;
        public string pathToFull;

        public MainWindow()                                                         //Главное окно
        {
            InitializeComponent();
            if (File.Exists("path.txt"))
            {
                using (StreamReader sr = File.OpenText("path.txt"))
                {
                    path = sr.ReadLine();
                }
                if (path == "")
                {
                    var dialog = new SaveFileDialog();
                    dialog.Title = "Выберите рабочую папку";
                    dialog.Filter = "Directory|*.this.directory";
                    dialog.FileName = "select";
                    if (dialog.ShowDialog() == true)
                    {
                        path = dialog.FileName;
                        path = path.Replace("\\select.this.directory", "");
                        path = path.Replace(".this.directory", "");
                        if (!Directory.Exists(path))
                        {
                            Directory.CreateDirectory(path);
                        }
                    }
                    using (StreamWriter sw = File.CreateText("path.txt"))
                    {
                        sw.WriteLine(path);
                    }
                }
            }
            else
            {
                var dialog = new SaveFileDialog();
                dialog.Title = "Select a Directory";
                dialog.Filter = "Directory|*.this.directory";
                dialog.FileName = "select";
                if (dialog.ShowDialog() == true)
                {
                    path = dialog.FileName;
                    path = path.Replace("\\select.this.directory", "");
                    path = path.Replace(".this.directory", "");
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                }
                using (StreamWriter sw = File.CreateText("path.txt"))
                {
                    sw.WriteLine(path);
                }
            }
            pathToShort = path + @"\Журнал отлова (short).xlsx";
            pathToFull = path + @"\Журнал отлова (full).xlsx";
            exApp = new Excel.Application();
            exWorkbook = exApp.Workbooks.Open(pathToShort);
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
