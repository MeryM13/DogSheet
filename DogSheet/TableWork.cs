using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DogSheet
{
    class TableWork
    {
        public Excel.Worksheet wsht;

        public TableWork(Excel.Worksheet wsht)
        {
            this.wsht = wsht;
        }

        public Excel.Range TableRangeSearch(string text)
        {

            Excel.Range currentFind;
            currentFind = wsht.get_Range("A2", GetLast()).Find(text);
            return currentFind;
        }

        public string[] GetRow(Excel.Range stringRng, string[] data)
        {
            int row = stringRng.Row;
            for (int i = 0; i < data.Length; i++)
            {
                var value = wsht.Cells[row, i + 1].Value2;
                if (value != null)
                {
                    if (value is double && i != 3)
                    {
                        DateTime dt = DateTime.FromOADate((double)value);
                        data[i] = dt.ToString("dd.MM.yyyy");
                    }
                    else
                    {
                        data[i] = value.ToString();
                    }
                }
                else
                {
                    data[i] = "";
                }
            }
            return data;
        }

        public void SetRow(Excel.Range stringRng, string[] data)
        {
            int row = stringRng.Row;
            for (int i = 0; i < data.Length; i++)
            {
                wsht.Cells[row, i + 1].Value2 = data[i];
            }
        }

        public Excel.Range GetLast()
        {
            return wsht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        }
    }
}