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

        //public bool TableSearch(string text)
        //{
        //    Excel.Range searchRng = wsht.get_Range("A1:A1000");
        //    Excel.Range currentFind = null;
        //    Excel.Range firstFind = null;
        //    currentFind = searchRng.Find(text);
        //    while (currentFind != null)
        //        if (firstFind == null)
        //        {
        //            firstFind = currentFind;
        //        }
        //        else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1) == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
        //        {
        //            return true;
        //        }
        //    return false;
        //}

        public Excel.Range TableRangeSearch(string text)
        {

            Excel.Range searchRng = wsht.get_Range("A2", GetLast());
            Excel.Range currentFind;
            currentFind = searchRng.Find(text);
            return currentFind;
        }

        public string[] GetRow(Excel.Range stringRng, string[] data)
        {
            int row = stringRng.Row;
            for (int i = 0; i < data.Length; i++)
            {
                if (wsht.Cells[row, i + 1].Value2 != null)
                    data[i] = wsht.Cells[row, i + 1].Value2.ToString();
                else
                    data[i] = "";
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