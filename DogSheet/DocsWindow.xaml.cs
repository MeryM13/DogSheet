using System;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace DogSheet
{
    public partial class DocsWindow : Window
    {
        public MainWindow MW;

        public Excel.Worksheet exWorksheet;

        public DocsWindow(MainWindow mainWindow)
        {
            InitializeComponent();

            MW = mainWindow;

            exWorksheet = MW.exWorkbook.Sheets[1];
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
