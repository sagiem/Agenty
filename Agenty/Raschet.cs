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
        public string file;


        public Raschet(string file)
        {
            this.file = file;
        }
        public void Exelreader()
        {
            string b,c;
            int d,e;
            int z = 14;

            //Excel.Range Rng;
            //Excel.Workbook xlWB;
            //Excel.Worksheet xlSht;
            //Excel.Application xlApp = new Excel.Application(); //создаём приложение Excel
            //xlWB = xlApp.Workbooks.Open(file); //открываем наш файл           
            //xlSht = xlWB.Worksheets["Лист1"]; //или так xlSht = xlWB.ActiveSheet //активный лист

            //a = xlSht.Cells[14, 3].Text.ToString();


            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(file); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист


            //var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            //string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
            //for (int i = 0; i < lastCell.Column; i++) //по всем колонкам
            //    for (int j = 0; j < lastCell.Row; j++) // по всем строкам
            b = ObjWorkSheet.Cells[z, 3].Text.ToString();
            while (b != "")
            {
                

                b = ObjWorkSheet.Cells[z, 3].Text.ToString();//считываем текст в строку
                c = ObjWorkSheet.Cells[z, 4].Text.ToString();
                if (b.Length == 10)
                {
                    d = int.Parse(b);
                    //e = int.Parse(c);
                }

                z++;

                //if(b=="")
                //{
                //    //ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
                //    //ObjWorkExcel.Quit(); // выйти из экселя
                //    //GC.Collect(); // убрать за  собой
                //    return;
                //}
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за  собой

        }


    }
}
