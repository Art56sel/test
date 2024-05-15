using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TESTVER0._1
{
    internal class ExcelHandler
    {
        public string savePathCIty = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, @"Resources\Ситилинк.xlsx");
        public string savePathOzon = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, @"Resources\Озон.xlsx"); 
        public string savePathYn = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, @"Resources\ЯндексМаркет.xlsx");
        public void CreateAndFillExcel()
        {
            Citylink();
            Ozon();
            YndexMarket();
        }

        #region создание таблицы Ситилинк 
        public void Citylink()
        {
            //Ситилинк
            if (!File.Exists(savePathCIty))
            {
                // Создаем новый объект Excel

                string filePathCIty = @"C:\Users\User\Downloads\citilink.xlsx";

                // Создаем новый объект Excel
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;

                // Открываем книгу citilink.xlsx
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePathCIty);
                Excel.Worksheet worksheet = workbook.Sheets[1];

                // Создаем новую книгу Excel
                Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet newWorksheet = newWorkbook.Sheets[1];

                // Копируем данные из citilink.xlsx
                for (int i = 1; i <= worksheet.UsedRange.Rows.Count; i++)
                {
                    newWorksheet.Cells[i, 1] = worksheet.Cells[i, 3].Value; // Ссылка
                    newWorksheet.Cells[i, 2] = worksheet.Cells[i, 5].Value; // Категория
                    newWorksheet.Cells[i, 3] = worksheet.Cells[i, 6].Value; // Арт
                    newWorksheet.Cells[i, 4] = worksheet.Cells[i, 7].Value; // Бренд
                    newWorksheet.Cells[i, 5] = worksheet.Cells[i, 8].Value; // Цена
                }

                // Сохраняем новую книгу Excel
                newWorkbook.SaveAs(savePathCIty);

                // Закрываем книги и приложение Excel
                newWorkbook.Close();
                workbook.Close();
                excelApp.Quit();


            }
        }
        #endregion
        #region создание таблицы Озон
        //Ozon
        public void Ozon()
        {
            if (!File.Exists(savePathOzon))
            {

                // Создаем новый объект Excel

                string filePathOZ = @"C:\Users\User\Downloads\OZON.xlsx";

                // Создаем новый объект Excel
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;

                // Открываем книгу citilink.xlsx
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePathOZ);
                Excel.Worksheet worksheet = workbook.Sheets[1];

                // Создаем новую книгу Excel
                Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet newWorksheet = newWorkbook.Sheets[1];

                // Копируем данные из citilink.xlsx
                for (int i = 1; i <= worksheet.UsedRange.Rows.Count; i++)
                {
                    newWorksheet.Cells[i, 1] = worksheet.Cells[i, 7].Value; // Ссылка
                    newWorksheet.Cells[i, 2] = worksheet.Cells[i, 6].Value; // Категория
                    newWorksheet.Cells[i, 3] = worksheet.Cells[i, 5].Value; // Арт
                    newWorksheet.Cells[i, 4] = worksheet.Cells[i, 9].Value; // Бренд
                    string price = worksheet.Cells[i, 8].Value?.ToString();
                    if (!string.IsNullOrEmpty(price))
                    {
                        newWorksheet.Cells[i, 5] = price.Replace("₽", "").Replace(" ", "");// Цена

                    }
                }

                // Сохраняем новую книгу Excel
                newWorkbook.SaveAs(savePathOzon);

                // Закрываем книги и приложение Excel
                newWorkbook.Close();
                workbook.Close();
                excelApp.Quit();

            }
        }
        #endregion
        #region создание таблицы ЯндексМаркет
        //ЯндексМаркет
        public void YndexMarket()
        {
          
            if (!File.Exists(savePathYn))
            {

                // Создаем новый объект Excel

                string filePath = @"C:\Users\User\Downloads\marketYandex.xlsx";

                // Создаем новый объект Excel
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;

                // Открываем книгу citilink.xlsx
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet worksheet = workbook.Sheets[1];

                // Создаем новую книгу Excel
                Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet newWorksheet = newWorkbook.Sheets[1];

                // Копируем данные из marketYandex.xlsx
                for (int i = 1; i <= worksheet.UsedRange.Rows.Count; i++)
                {
                    newWorksheet.Cells[i, 1] = worksheet.Cells[i, 5].Value; // наименование
                    newWorksheet.Cells[i, 2] = worksheet.Cells[i, 6].Value; // Категория
                    newWorksheet.Cells[i, 3] = worksheet.Cells[i, 9].Value; // Арт
                    newWorksheet.Cells[i, 4] = worksheet.Cells[i, 7].Value; // Бренд
                   
                    string price = worksheet.Cells[i, 8].Value?.ToString();
                    if (!string.IsNullOrEmpty(price))
                    {
                        string x = price.Replace("₽", "");
                        price = "";
                        price = x;
                        newWorksheet.Cells[i, 5] = price.Replace("Цена с картой Яндекс Пэй:", "").Replace(" ", "");
                    }
                }

                // Сохраняем новую книгу Excel
                newWorkbook.SaveAs(savePathYn);

                // Закрываем книги и приложение Excel
                newWorkbook.Close();
                workbook.Close();
                excelApp.Quit();
            }
        }
        #endregion

    }


}

