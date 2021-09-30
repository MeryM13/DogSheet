using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace DogSheet
{
    class DocsWork
    {
        public Word.Application wordApp = new Word.Application();
        private object missing = System.Reflection.Missing.Value;

        public void Doc1Create(string photoPath, string animalNumber, string catchDate, string catchPlace, string requestNumber, string requestDate,
            string head, string catcher, string category, string type, string sex, string breed, string color, string fur, string ears, string tail, string age,
            string weight, string additional, string chip, string awayDate)
        {
            object filename = @"C:\Users\pshar\Desktop\Карточка учета " + animalNumber + ".docx";

            Word.Document doc;
            try
            {
                doc = wordApp.Documents.Open(filename);                                            //Попытка открыть существующий файл, если его нет, создание нового
            }
            catch
            {
                doc = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            }
            Word.Range rng = doc.Range(doc.Content.Start, doc.Content.End);                       //Очистка документа
            rng.Text = "";
            doc.Content.SetRange(0, 0);
            Word.Paragraph paragraph1 = doc.Content.Paragraphs.Add(ref missing);                    //Создание первой таблицы
            Word.Table table1 = doc.Tables.Add(paragraph1.Range, 1, 2, ref missing, ref missing);
            table1.Borders.Enable = 1;                                                               //Заполнение первой таблицы
            foreach (Word.Row row in table1.Rows)
            {
                foreach (Word.Cell cell in row.Cells)
                { 
                    cell.Range.Font.Name = "Times New Roman";
                    cell.Range.Font.Size = 12;
                    cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }
            }
            table1.Cell(1, 1).Range.InlineShapes.AddPicture(photoPath);
            table1.Cell(1, 2).Range.Text = "КАРТОЧКА\nУЧЕТА БЕЗНАДЗОРНОГО ЖИВОТНОГО\n№ " + animalNumber + "\n\n" 
                + catchDate + "\n" +catchPlace;
            paragraph1.Range.InsertParagraphAfter();
            Word.Paragraph paragraph2 = doc.Content.Paragraphs.Add(ref missing);
            paragraph2.Range.Font.Name = "Times New Roman";
            paragraph2.Range.Font.Size = 12;
            paragraph2.Range.Text = "1.	В соответствии с заявлением № " + requestNumber + " от " + requestDate + " Служба помощи животным «Белый Пёс» " +
                "(ИП Шаромова А.Ю. ОГРНИП 319547600032822) в составе: руководитель " + head + " и ловец " + catcher + " на машине: " +
                "MITSUBISHI DELICA D3, госномер Н472НО 154 произвела отлов и транспортировку животного:	";
            paragraph2.Range.InsertParagraphAfter();
            Word.Table table2 = doc.Tables.Add(paragraph2.Range, 15, 2, ref missing, ref missing);
            table2.Borders.Enable = 1;
            foreach (Word.Row row in table2.Rows)
            {
                foreach (Word.Cell cell in row.Cells)
                {
                    cell.Range.Font.Name = "Times New Roman";
                    cell.Range.Font.Size = 12;
                }
            }
            table2.Cell(1, 1).Range.Text = "Категория: собака, щенок, кошка, котенок, иное";
            table2.Cell(1, 2).Range.Text = category;
            table2.Cell(2, 1).Range.Text = "Дата поступления в организацию по отлову безнадзорных животных";
            table2.Cell(2, 2).Range.Text = catchDate;
            table2.Cell(3, 1).Range.Text = "Вид";
            table2.Cell(3, 2).Range.Text = type;
            table2.Cell(4, 1).Range.Text = "Пол";
            table2.Cell(4, 2).Range.Text = sex;
            table2.Cell(5, 1).Range.Text = "Порода";
            table2.Cell(5, 2).Range.Text = breed;
            table2.Cell(6, 1).Range.Text = "Окрас";
            table2.Cell(6, 2).Range.Text = color;
            table2.Cell(7, 1).Range.Text = "Шерсть";
            table2.Cell(7, 2).Range.Text = fur;
            table2.Cell(8, 1).Range.Text = "Уши";
            table2.Cell(8, 2).Range.Text = ears;
            table2.Cell(9, 1).Range.Text = "Хвост";
            table2.Cell(9, 2).Range.Text = tail;
            table2.Cell(10, 1).Range.Text = "Вес";
            table2.Cell(10, 2).Range.Text = weight;
            table2.Cell(11, 1).Range.Text = "Возраст (примерный)";
            table2.Cell(11, 2).Range.Text = age;
            table2.Cell(12, 1).Range.Text = "Особые приметы";
            table2.Cell(12, 2).Range.Text = additional;
            table2.Cell(13, 1).Range.Text = "Идентификационная метка, чип (способ и место нанесения)";
            table2.Cell(13, 2).Range.Text = chip;
            table2.Cell(14, 1).Range.Text = "Регистрационный номер";
            table2.Cell(14, 2).Range.Text = animalNumber;
            table2.Cell(15, 1).Range.Text = "Место отлова (адрес)";
            table2.Cell(15, 2).Range.Text = catchPlace;
            paragraph2.Range.InsertParagraphAfter();
            Word.Paragraph paragraph3 = doc.Content.Paragraphs.Add(ref missing);
            paragraph3.Range.Font.Name = "Times New Roman";
            paragraph3.Range.Font.Size = 12;
            paragraph3.Range.Text = "2.	Осуществлена передача безнадзорного животного владельцу, в организацию, возврат на " +
                "прежнее место обитания. Дата: " + awayDate + "\nДанные: для юридических лиц:\nорганизация" +
                "__________________________________________________________________,адрес" +
                "________________________________________________________________________,телефон" +
                "______________________________________________________________________,Ф.И.О.руководителя" +
                "___________________________________________________________,Ф.И.О.и телефон ответственного " +
                "за содержание(если он есть)________________________________________________________________________" +
                ";\n            для физических лиц: Ф.И.О._______________________________________________________________________," +
                "адрес________________________________________________________________________,телефон" +
                "______________________________________________________________________,паспортные данные" +
                "_______________________________________________________________________.\n\nДата выписки животного " +
                "_____________________________________________________________________________\nФ.И.О. руководителя " +
                "___________________________________________\nПодпись__________________________\n";
            paragraph3.Range.InsertParagraphAfter();
            Word.Paragraph paragraph4 = doc.Content.Paragraphs.Add(ref missing);
            paragraph4.Range.Font.Name = "Times New Roman";
            paragraph4.Range.Font.Size = 12;
            paragraph4.Range.Text = "3. Оформление в муниципальную собственность\nДата / " +
                "номер документа_________________________________________________________\n";
            paragraph4.Range.InsertParagraphAfter();
            doc.SaveAs2(ref filename);
            doc.Close(ref missing, ref missing, ref missing);
        }

        public void Doc2Create(string photoPath, string animalNumber, string catchDate, string sex, string breed, string color, 
            string fur, string age, string weight, string additional, string medical, string awayDate, string away,
            string vaccine, string vacDate, string label, string mark, string stMethod)
        {
            object filename = @"C:\Users\pshar\Desktop\Вет карта " + animalNumber + ".docx";

            Word.Document doc;
            try
            {
                doc = wordApp.Documents.Open(filename);                                            //Попытка открыть существующий файл, если его нет, создание нового
            }
            catch
            {
                doc = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            }
            Word.Range rng = doc.Range(doc.Content.Start, doc.Content.End);                       //Очистка документа
            rng.Text = "";
            doc.Content.SetRange(0, 0);

            Word.Paragraph para1 = doc.Content.Paragraphs.Add(ref missing);

            Word.Table table1 = doc.Tables.Add(para1.Range, 13, 2, ref missing, ref missing);

            table1.Borders.Enable = 1;
            foreach (Word.Row row in table1.Rows)
            {
                foreach (Word.Cell cell in row.Cells)
                {
                    cell.Range.Font.Name = "Times New Roman";
                    cell.Range.Font.Size = 14;
                    cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }
            }
            table1.Cell(1, 1).Range.InlineShapes.AddPicture(photoPath);
            table1.Cell(1, 2).Range.Text = "УЧЕТНАЯ КАРТОЧКА ЖИВОТНОГО № " + animalNumber;
            table1.Cell(1, 2).Range.Font.Size = 16;
            table1.Cell(1, 2).Range.Font.Bold = 1;
            table1.Cell(2, 1).Range.Text = "Дата поступления животного в приют";
            table1.Cell(2, 2).Range.Text = catchDate;
            table1.Cell(3, 1).Range.Text = "Вид (порода)";
            table1.Cell(3, 2).Range.Text = breed;
            table1.Cell(4, 1).Range.Text = "Описание";
            table1.Cell(4, 2).Range.Text = color + ", " + fur;
            table1.Cell(5, 1).Range.Text = "Вес";
            table1.Cell(5, 2).Range.Text = weight;
            table1.Cell(6, 1).Range.Text = "Возраст (примерный)";
            table1.Cell(6, 2).Range.Text = age;
            table1.Cell(7, 1).Range.Text = "Пол";
            table1.Cell(7, 2).Range.Text = sex;
            table1.Cell(8, 1).Range.Text = "Особые приметы";
            table1.Cell(8, 2).Range.Text = additional;
            table1.Cell(9, 1).Range.Text = "Осмотр";
            table1.Cell(9, 2).Range.Text = medical;
            table1.Cell(10, 1).Range.Text = "Вакцинация животного";
            table1.Cell(10, 2).Split(2, 2);
            table1.Cell(10, 2).Range.Text = "Вакцина";
            table1.Cell(10, 3).Range.Text = "Дата";
            table1.Cell(11, 2).Range.Text = vaccine;
            table1.Cell(11, 3).Range.Text = vacDate;
            table1.Cell(12, 1).Range.Text = "Данные о маркировании животного";
            table1.Cell(12, 2).Split(2, 2);
            table1.Cell(12, 2).Range.Text = "Бирка";
            table1.Cell(12, 3).Range.Text = "Клеймо";
            table1.Cell(13, 2).Range.Text = label;
            table1.Cell(13, 3).Range.Text = mark;
            table1.Cell(14, 1).Range.Text = "Стерилизация животного";
            table1.Cell(14, 2).Range.Text = stMethod;
            table1.Cell(15, 1).Range.Text = "Выбытие";
            table1.Cell(15, 2).Range.Text = awayDate + ", " + away;

            doc.SaveAs2(ref filename);
            doc.Close(ref missing, ref missing, ref missing);
        }
    }
}
