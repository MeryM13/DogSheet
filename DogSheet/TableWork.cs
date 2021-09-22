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
        private Excel.Worksheet wsht;
        public TableWork(Excel.Worksheet worksheet)
        {
            //TableWork TW = new TableWork();
            wsht = worksheet;
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
            Excel.Range searchRng = wsht.get_Range("A2:A1000");
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            currentFind = searchRng.Find(text);
            //while (currentFind != null)
            //    if (firstFind == null)
            //    {
            //        firstFind = currentFind;
            //    }
            //    else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1) == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
            //    {
            //        return currentFind;
            //    }
            //return null;
            return currentFind;
        }

        //public void GetRow(Excel.Range stringRng, string number, string curator, string phone, string catchPlace, string type, string color,
        //    string additional, string pregnant, string trauma, string stPlace, string stDate, string mark, string label, string vaccine,
        //    string vacDate, string away, string AwayDate)
        //{
        //    int row = stringRng.Row;
        //    number = wsht.Cells[row, 1].Value.ToString();
        //    curator = wsht.Cells[row, 2].Value.ToString();
        //    phone = wsht.Cells[row, 3].Value.ToString();
        //    catchPlace = wsht.Cells[row, 4].Value.ToString();
        //    type = wsht.Cells[row, 5].Value.ToString();
        //    color = wsht.Cells[row, 6].Value.ToString();
        //    additional = wsht.Cells[row, 7].Value.ToString();
        //    pregnant = wsht.Cells[row, 8].Value.ToString();
        //    trauma = wsht.Cells[row, 9].Value.ToString();
        //    stPlace = wsht.Cells[row, 10].Value.ToString();
        //    stDate = wsht.Cells[row, 11].Value.ToString();
        //    mark = wsht.Cells[row, 12].Value.ToString();
        //    label = wsht.Cells[row, 13].Value.ToString();
        //    vaccine = wsht.Cells[row, 14].Value.ToString();
        //    vacDate = wsht.Cells[row, 15].Value.ToString();
        //    away = wsht.Cells[row, 16].Value.ToString(); ;
        //    AwayDate = wsht.Cells[row, 17].Value.ToString();
        //}

        public void SetRow(Excel.Range stringRng, string number, string curator, string phone, string catchPlace, string type, string color,
            string additional, string pregnant, string trauma, string stPlace, string stDate, string mark, string label, string vaccine,
            string vacDate, string away, string AwayDate)
        {

        }
    }
}
