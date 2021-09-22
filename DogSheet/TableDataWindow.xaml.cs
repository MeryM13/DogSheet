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
using Word = Microsoft.Office.Interop.Word;

namespace DogSheet
{
    /// <summary>
    /// Логика взаимодействия для TableDataWindow.xaml
    /// </summary>
    public partial class TableDataWindow : Window
    {
        Excel.Application app = new Excel.Application();
        Excel.Workbook workbook;
        AddDataWindow ADW = new AddDataWindow();
        

        public TableDataWindow()
        {
            InitializeComponent();
            //if (workbook == null)
            //    workbook = app.Workbooks.Open(@"C:\Users\pshar\source\repos\DogSheet\Журнал отлова безнадзорных животных.xlsx");
        }

        public TableDataWindow(Excel.Range startRng)
        {
            workbook = app.Workbooks.Open(@"C:\Users\pshar\source\repos\DogSheet\Журнал отлова безнадзорных животных.xlsx");
            Excel.Worksheet worksheet = workbook.Sheets[1];
            //TableWork TW = new TableWork(worksheet);
            //TW.GetRow(startRng, NumberTextbox.Text, CuratorTextbox.Text, PhoneTextbox.Text, CatchPlaceTextbox.Text, TypeTextbox.Text, ColorTextbox.Text,
            //AdditionalTextbox.Text, PregnantTextbox.Text, TraumaTextbox.Text, StPlaceTextbox.Text, StDateTextbox.Text, MarkTextbox.Text, LabelTextbox.Text, 
            //VacTextbox.Text, VacDateTextbox.Text, AwayTextbox.Text, AwayDateTextbox.Text);
            int row = startRng.Row;
            NumberTextbox.Text = worksheet.Cells[row, 1].Value2.ToString();
            CuratorTextbox.Text = worksheet.Cells[row, 2].Value2.ToString();
            PhoneTextbox.Text = worksheet.Cells[row, 3].Value2.ToString();
            CatchPlaceTextbox.Text = worksheet.Cells[row, 4].Value2.ToString();
            TypeTextbox.Text = worksheet.Cells[row, 5].Value2.ToString();
            ColorTextbox.Text = worksheet.Cells[row, 6].Value2.ToString();
            AdditionalTextbox.Text = worksheet.Cells[row, 7].Value2.ToString();
            PregnantTextbox.Text = worksheet.Cells[row, 8].Value2.ToString();
            TraumaTextbox.Text = worksheet.Cells[row, 9].Value2.ToString();
            StPlaceTextbox.Text = worksheet.Cells[row, 10].Value2.ToString();
            StDateTextbox.Text = worksheet.Cells[row, 11].Value2.ToString();
            MarkTextbox.Text = worksheet.Cells[row, 12].Value2.ToString();
            LabelTextbox.Text = worksheet.Cells[row, 13].Value2.ToString();
            VacTextbox.Text = worksheet.Cells[row, 14].Value2.ToString();
            VacDateTextbox.Text = worksheet.Cells[row, 15].Value2.ToString();
            AwayTextbox.Text = worksheet.Cells[row, 16].Value2.ToString(); ;
            AwayDateTextbox.Text = worksheet.Cells[row, 17].Value2.ToString();
            Console.WriteLine(NumberTextbox.Text, CuratorTextbox.Text, PhoneTextbox.Text, CatchPlaceTextbox.Text, TypeTextbox.Text, ColorTextbox.Text,
            AdditionalTextbox.Text, PregnantTextbox.Text, TraumaTextbox.Text, StPlaceTextbox.Text, StDateTextbox.Text, MarkTextbox.Text, LabelTextbox.Text, 
            VacTextbox.Text, VacDateTextbox.Text, AwayTextbox.Text, AwayDateTextbox.Text);
            InitializeComponent();
        }

        private void ForwardButton_Click(object sender, RoutedEventArgs e)
        {
            ADW.Show();
            this.Close();
            workbook.Close();
        }
    }
}
