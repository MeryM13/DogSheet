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
        Excel.Application DWapp;
        Excel.Workbook DWworkbook;
        Excel.Worksheet DWworksheet;

        public DocsWindow(Excel.Application app, Excel.Workbook workbook)
        {
            DWapp = app;
            DWworkbook = workbook;
            DWworksheet = DWworkbook.Sheets[1];
            InitializeComponent();
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            //DWworkbook = DWapp.Workbooks.Open(@"C:\Users\pshar\source\repos\DogSheet\Журнал отлова безнадзорных животных.xlsx");
            TableWork TW = new TableWork(DWworksheet);
            Excel.Range rng = null;
            if (NumberTextbox.Text == "")                               //Создание новой записи в таблице
            {
                rng = TW.GetLast();
                TableDataWindow TDW = new TableDataWindow(DWapp, DWworkbook, DWworksheet, rng);
                //DWworkbook.Close();
                TDW.Show();
                this.Close();
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
                    //workbook.Close();
                    TableDataWindow TDW = new TableDataWindow(DWapp, DWworkbook, DWworksheet, rng);
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
