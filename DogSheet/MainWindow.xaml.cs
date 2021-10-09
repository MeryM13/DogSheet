using System.IO;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;


namespace DogSheet
{
    public partial class MainWindow : Window    //Главное окно
    {
        public Excel.Application exApp;
        public Excel.Workbook exWorkbook;

        public string path;
        public string pathToShort;
        public string pathToFull;

        public MainWindow()
        {
            InitializeComponent();
            if (File.Exists("path.txt"))            //попытка считать пути из файла
            {
                using (StreamReader sr = File.OpenText("path.txt")) //если файл есть
                {
                    path = sr.ReadLine();             //считать значения
                    pathToShort = sr.ReadLine();
                    pathToFull = sr.ReadLine();
                }
                if (path == "")                  //если нет пути для рабочей папки
                {
                    DirectoryChoice();
                }
                if (pathToShort == "")          //если нет пути для краткой таблицы
                {
                    ShortChoice();
                }
                if (pathToFull == "")           //если нет пути для полной таблицы
                {
                    FullChoice();
                }
            }
            else        //если файла с путями нет
            {
                DirectoryChoice();
                ShortChoice();
                FullChoice();
            }

            exApp = new Excel.Application();    //запуск Excel

            bool opened;                        //проверка правильности пути до таблицы
            do
            {
                opened = true;
                try
                {
                    exWorkbook = exApp.Workbooks.Open(pathToShort); //открыть краткую таблицу
                }
                catch           //если был указан неверный путь
                {
                    ShortChoice();      //открыть диалог для нового назначения пути
                    opened = false;
                }
            } while (!opened);          //пока таблица не откроется
        }

        private void TablesButton_Click(object sender, RoutedEventArgs e)           //открытие таблиц для просмотра
        {
            TablesWindow TW = new TablesWindow(this);

            TW.Show();
            this.Hide();
        }

        private void DocsButton_Click(object sender, RoutedEventArgs e)             //открытие окна для создания отчетов
        {
            DocsWindow DW = new DocsWindow(this);

            DW.Show();
            this.Hide();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) //закрытие окна
        {
            exWorkbook.Close(true);
            exApp.Quit();
        }

        private void Button_Click(object sender, RoutedEventArgs e)     //Нажатие по кнопке смены папки
        {
            DirectoryChoice();
        }

        private void DirectoryChoice()              //диалог для выбора папки
        {
            SaveFileDialog dialog = new SaveFileDialog
            {
                Title = "Выберите рабочую папку",
                Filter = "Directory|*.this.directory",
                FileName = "select"
            };
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
            WriteInFile();
        }

        public void ShortChoice()           //диалог для выбора краткой таблицы
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Выберите краткую таблицу",
                Filter = "Excel files(*.xlsx)|*.xlsx| All files(*.*) | *.*"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                pathToShort = openFileDialog.FileName;
            }
            WriteInFile();
        }

        public void FullChoice()            //диалог для выбора полной таблицы
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Выберите полную таблицу",
                Filter = "Excel files(*.xlsx)|*.xlsx| All files(*.*) | *.*"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                pathToFull = openFileDialog.FileName;
            }
            WriteInFile();
        }

        private void WriteInFile()      //запись путей в файл
        {
            using (StreamWriter sw = File.CreateText("path.txt"))
            {
                sw.WriteLine(path);
                sw.WriteLine(pathToShort);
                sw.WriteLine(pathToFull);
            }
        }
    }
}
