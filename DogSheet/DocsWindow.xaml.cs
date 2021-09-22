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
using Excel = Microsoft.Office.Interop.Excel;

namespace DogSheet
{
    /// <summary>
    /// Логика взаимодействия для DocsWindow.xaml
    /// </summary>
    public partial class DocsWindow : Window
    {
        public TableDataWindow TDW = new TableDataWindow();
        Excel.Application app = new Excel.Application();
        Excel.Workbook workbook;

        public DocsWindow()
        {
            InitializeComponent();
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            if (NumberTextbox.Text == "")                               //Создание новой записи в таблице
            {
                TableDataWindow TDW = new TableDataWindow();
                TDW.Show();
                this.Close();
            }
            else                                                        //Работа с уже имеющейся записью
            {
                workbook = app.Workbooks.Open(@"C:\Users\pshar\source\repos\DogSheet\Журнал отлова безнадзорных животных.xlsx");
                Excel.Worksheet worksheet = (Excel.Worksheet)app.ActiveWorkbook.Sheets[1];
                TableWork TW = new TableWork(worksheet);
                Excel.Range rng = null;
                Console.WriteLine(NumberTextbox.Text);
                if (NumberTextbox.Text.Length <= 3)                     //Если введены только цифры
                {
                    if (TW.TableRangeSearch(NumberTextbox.Text + " БП") != null)
                    {
                        rng = TW.TableRangeSearch(NumberTextbox.Text + " БП");
                    }
                }
                else                                                    //Если введено с индексом БП
                {
                    if (TW.TableRangeSearch(NumberTextbox.Text) != null)
                    {
                        rng = TW.TableRangeSearch(NumberTextbox.Text);
                    }
                }
                if (rng != null)
                {
                    workbook.Close();
                    TableDataWindow TDW = new TableDataWindow();
                    TDW.Show();
                    this.Close();
                }
                else                                                    //Если нет такого номера или введены неправилные данные
                {
                    MessageBox.Show("Животного под таким номером не существует, попробуйте ещё раз.");
                }
            }
        }
    }
}
