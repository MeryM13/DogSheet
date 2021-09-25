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
        private Excel.Application TDWapp;
        private Excel.Workbook TDWworkbook;
        private Excel.Worksheet TDWworksheet;
        private string[] data = new string[18];
        private AddDataWindow ADW = new AddDataWindow();
        
        public TableDataWindow(Excel.Application app, Excel.Workbook workbook, Excel.Worksheet worksheet, Excel.Range startRng)
        {
            InitializeComponent();
            TDWapp = app;
            TDWworkbook = workbook;
            TDWworksheet = worksheet;
            TableWork TW = new TableWork(TDWworksheet);
            TW.GetRow(startRng, data);
            NumberTextbox.Text = data[0];
            CatchDateTextbox.Text = data[1];
            CuratorTextbox.Text = data[2];
            PhoneTextbox.Text = data[3];
            CatchPlaceTextbox.Text = data[4];
            TypeTextbox.Text = data[5];
            ColorTextbox.Text = data[6];
            AdditionalTextbox.Text = data[7];
            PregnantTextbox.Text = data[8];
            TraumaTextbox.Text = data[9];
            StPlaceTextbox.Text = data[10];
            StDateTextbox.Text = data[11];
            MarkTextbox.Text = data[12];
            LabelTextbox.Text = data[13];
            VacTextbox.Text = data[14];
            VacDateTextbox.Text = data[15];
            AwayTextbox.Text = data[16];
            AwayDateTextbox.Text = data[17];
        }

        private void ForwardButton_Click(object sender, RoutedEventArgs e)
        {
            ADW.Show();
            this.Close();
            //workbook.Close();
        }
    }
}
