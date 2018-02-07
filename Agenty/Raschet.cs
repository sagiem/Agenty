using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Agenty
{
    class Raschet
    {
        public string File;
       

        public Raschet(string file)
        {
            this.File=file;
        }
        public void Exelreader()
        {
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(File); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
            for (int i = 0; i < lastCell.Column; i++) //по всем колонкам
                for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                    list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой
        }


    }
}
