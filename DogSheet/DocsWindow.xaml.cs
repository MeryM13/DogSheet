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
        public Excel.Worksheet exWorksheet;
        public MainWindow MW;

        public DocsWindow(MainWindow mainWindow)
        {
            MW = mainWindow;
            exWorksheet = MW.exWorkbook.Sheets[1];
            InitializeComponent();
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            TableWork TW = new TableWork(exWorksheet);
            Excel.Range rng = null;
            if (NumberTextbox.Text == "")                               //Создание новой записи в таблице
            {
                Excel.Range last = TW.GetLast();
                rng = exWorksheet.Cells[last.Row + 1, 1];
                TableDataWindow TDW = new TableDataWindow(rng, this);
                TDW.Show();
                this.Hide();
            }
            else                                                        //Работа с уже имеющейся записью
            {
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
                    TableDataWindow TDW = new TableDataWindow(rng, this);
                    TDW.Show();
                    this.Hide();
                }
                else                                                    //Если нет такого номера или введены неправилные данные
                {
                    MessageBox.Show("Животного под таким номером не существует, попробуйте ещё раз.");
                }
            }
        }

        private void Window_Closed(object sender, EventArgs e)          //При закрытии окна снова откроется главное
        {
            MW.Show();
        }
    }
}
