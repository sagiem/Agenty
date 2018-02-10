using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Agenty
{
    class Raschet
    {
        public string file;
        string b, c, e;
        decimal d;
        int z = 14;
        int t = 1;

        List<ExcelOpen> exp = new List<ExcelOpen>();


        public Raschet(string file)
        {
            this.file = file;
        }
        public void Exelreader()
        {
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(file); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            b = ObjWorkSheet.Cells[z, 3].Text.ToString();
            while (b != "")
            {
                b = ObjWorkSheet.Cells[z, 3].Text.ToString();//считываем текст в строку
                c = ObjWorkSheet.Cells[z, 4].Text.ToString();
                if (b.Length == 10)
                {
                    e = b;
                    d =decimal.Parse(c.Replace(" ", ""));
                    ExcelOpen excelOpen = new ExcelOpen(t,e,d);
                    exp.Add(excelOpen);
                    t++;
                }
                z++;               
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за  собой

        }

        public void ExelAkt()
        {

            //Создаём новый Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //Загружаем документ
            Microsoft.Office.Interop.Word.Document doc = null;

            object fileName = @"C:\Новая папка\Desktop\прога\Agenty\files\akt.doc";
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;

            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            //Указываем таблицу в которую будем помещать данные (таблица должна существовать в шаблоне документа!)
            Microsoft.Office.Interop.Word.Table tbl = app.ActiveDocument.Tables[1];

            //Заполняем в таблицу - 10 записей.

            for (int i = 1; i < 11; i++)
            {
                tbl.Rows.Add(ref missing);//Добавляем в таблицу строку.
                                          //Обычно саздаю только строку с заголовками и одну пустую для данных.
                tbl.Rows[i + 1].Cells[1].Range.Text = "Запись №" + i.ToString();
                tbl.Rows[i + 1].Cells[2].Range.Text = "Запись №" + i.ToString();
                tbl.Rows[i + 1].Cells[3].Range.Text = "Запись №" + i.ToString();

            }

            //Открываем документ для просмотра.
            app.Visible = true;


            //// Создаём экземпляр нашего приложения
            //Excel.Application excelApp = new Excel.Application();
            //// Создаём экземпляр рабочий книги Excel
            //Excel.Workbook workBook;
            //// Создаём экземпляр листа Excel
            //Excel.Worksheet workSheet;

            //workBook = excelApp.Workbooks.Add();
            //workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

            //// Заполняем первую строку числами от 1 до 10
            //for (int j = 1; j <= exp.Count(); j++)
            //{
            //    workSheet.Cells[j, 1] = (exp[j-1]).a;
            //    workSheet.Cells[j, 2] = (exp[j - 1]).c;
            //    workSheet.Cells[j, 3] = (exp[j - 1]).d;
            //}

            //// Открываем созданный excel-файл
            //excelApp.Visible = true;
            //excelApp.UserControl = true;



        }


    }
}
