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
    public partial class TableDataWindow : Window                   //Окно для ввода данных, которые есть в первой таблице
    {
        private Excel.Range workRange;
        public string[] data = new string[18];
        private TableWork TW;
        public DocsWindow DW;
        private AddDataWindow ADW;

        public TableDataWindow(Excel.Range startRng, DocsWindow docsWindow)
        {
            InitializeComponent();
            DW = docsWindow;
            workRange = startRng;
            TW = new TableWork(DW.exWorksheet);
            TW.GetRow(workRange, data);
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
            AwayDateTextbox.Text = data[16];
            AwayTextbox.Text = data[17];
        }

        private void ForwardButton_Click(object sender, RoutedEventArgs e)
        {
            data[0] = NumberTextbox.Text;
            data[1] = CatchDateTextbox.Text;
            data[2] = CuratorTextbox.Text;
            data[3] = PhoneTextbox.Text;
            data[4] = CatchPlaceTextbox.Text;
            data[5] = TypeTextbox.Text;
            data[6] = ColorTextbox.Text;
            data[7] = AdditionalTextbox.Text;
            data[8] = PregnantTextbox.Text;
            data[9] = TraumaTextbox.Text;
            data[10] = StPlaceTextbox.Text;
            data[11] = StDateTextbox.Text;
            data[12] = MarkTextbox.Text;
            data[13] = LabelTextbox.Text;
            data[14] = VacTextbox.Text;
            data[15] = VacDateTextbox.Text;
            data[16] = AwayDateTextbox.Text;
            data[17] = AwayTextbox.Text;
            TW.SetRow(workRange, data);
            ADW = new AddDataWindow(this);
            ADW.Show();
            this.Hide();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            DW.Show();
        }
    }
}
