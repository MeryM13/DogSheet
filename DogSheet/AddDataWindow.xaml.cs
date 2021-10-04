using System;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace DogSheet
{
    public partial class AddDataWindow : Window
    {
        private TableDataWindow TDW;
        private TableWork TW;

        private Excel.Workbook workbookFull;
        private Excel.Worksheet worksheetFull;
        private Excel.Range rng;

        private string[] allData = new string[33];
        private string photoPath = "";

        public AddDataWindow(TableDataWindow tableDataWindow)
        {
            InitializeComponent();

            TDW = tableDataWindow;

            workbookFull = TDW.DW.MW.exApp.Workbooks.Open(TDW.DW.MW.pathToFull);
            worksheetFull = workbookFull.Sheets[1];

            TW = new TableWork(worksheetFull);

            allData[0] = TDW.data[0];

            if (TW.TableRangeSearch(allData[0]) != null)
            {
                rng = TW.TableRangeSearch(allData[0]);

                TW.SetRow(rng, TDW.data);
                TW.GetRow(rng, allData);

                Doc1Checkbox.IsChecked = true;
                Doc2Checkbox.IsChecked = true;

                RequestNumberTextbox.Text = allData[18];
                RequestDateTextbox.Text = allData[19];
                HeadTextbox.Text = allData[20];

                if (allData[20] == allData[21])
                    GroupCheckbox.IsChecked = true;
                else
                    CatcherTextbox.Text = allData[21];


                if (allData[22] != null)
                    CategoryCombobox.SelectedItem = allData[22];


                if (allData[23] != null)
                    SexCombobox.SelectedItem = allData[23];


                BreedTextbox.Text = allData[24];

                if (allData[25] != null)
                    FurCombobox.SelectedItem = allData[25];

                if (allData[26] != null)
                    EarsCombobox.SelectedItem = allData[26];

                if (allData[27] != null)
                    TailCombobox.SelectedItem = allData[27];

                WeightTextbox.Text = allData[28];
                AgeTextbox.Text = allData[29];
                ChipTextbox.Text = allData[30];
                MedicalTextbox.Text = allData[31];
                StMethodTextbox.Text = allData[32];
            }
            else
            {
                Excel.Range search = TW.GetLast();
                rng = worksheetFull.Cells[search.Row + 1, 1];

                TW.SetRow(rng, TDW.data);
            }
        }

        private void PhotoButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                photoPath = openFileDialog.FileName;
            }
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            allData[18] = RequestNumberTextbox.Text;
            allData[19] = RequestDateTextbox.Text;
            allData[20] = HeadTextbox.Text;
            allData[21] = CatcherTextbox.Text;

            if (CategoryCombobox.SelectedItem == null)
                allData[22] = "";
            else
                allData[22] = CategoryCombobox.SelectedItem.ToString();

            if (SexCombobox.SelectedItem == null)
                allData[23] = "";
            else
                allData[23] = SexCombobox.SelectedItem.ToString();

            allData[24] = BreedTextbox.Text;

            if (FurCombobox.SelectedItem == null)
                allData[25] = "";
            else
                allData[25] = FurCombobox.SelectedItem.ToString();

            if (EarsCombobox.SelectedItem == null)
                allData[26] = "";
            else
                allData[26] = EarsCombobox.SelectedItem.ToString();

            if (TailCombobox.SelectedItem == null)
                allData[27] = "";
            else
                allData[27] = TailCombobox.SelectedItem.ToString();

            allData[28] = WeightTextbox.Text;
            allData[29] = AgeTextbox.Text;
            allData[30] = ChipTextbox.Text;
            allData[31] = MedicalTextbox.Text;
            allData[32] = StMethodTextbox.Text;

            TW.SetRow(rng, allData);

            if (photoPath != "")
            {
                DocsWork docsWork = new DocsWork();

                if (Doc1Checkbox.IsChecked == true)
                {
                    string additional = "";

                    if (allData[7] != "")
                        additional += allData[7];
                    if (allData[8] != "")
                        if (additional != "")
                            additional += ", " + allData[8];
                        else
                            additional += allData[8];
                    if (allData[9] != "")
                        if (additional != "")
                            additional += ", " + allData[9];
                        else
                            additional += allData[9];

                    docsWork.Doc1Create(TDW.DW.MW.path, photoPath, allData[0], allData[1], allData[4], allData[18], allData[19], allData[20], allData[21],
                        allData[22], allData[5], allData[23], allData[24], allData[6], allData[25], allData[26], allData[27], allData[29],
                        allData[28], additional, allData[30], allData[16]);
                }

                if (Doc2Checkbox.IsChecked == true)
                {
                    string sex = allData[23];

                    switch (allData[22])
                        {
                        case "Собака":
                        case "Щенок":
                            {
                                if (sex == "м")
                                    sex = "Кобель";
                                else
                                    sex = "Сука";
                                break;
                            }
                        case "Кошка":
                        case "Котенок":
                            {
                                if (sex == "м")
                                    sex = "Кот";
                                else
                                    sex = "Кошка";
                                break;
                            }
                        default:
                            {
                                sex = allData[23];
                                break;
                            }
                    }

                    string away;

                    if (allData[17] != "")
                        away = allData[17];
                    else
                        away = "выпуск";

                    docsWork.Doc2Create(TDW.DW.MW.path, photoPath, allData[0], allData[1], sex, allData[24], allData[6], allData[25], allData[29], allData[28], allData[7],
                        allData[31], allData[16], away, allData[14], allData[15], allData[13], allData[12], allData[32]);
                }

                docsWork.wordApp.Quit();
                Close();
                TDW.Close();
            }
            else
            {
                _ = MessageBox.Show("Вы не выбрали фотографию");
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            workbookFull.Close(true);
            TDW.Show();
        }
    }
}
