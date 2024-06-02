using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Khalin_Kypcova_612pst.Classes;
namespace APZ_lab3_khalin
{
    public static class CreateExcelList
    {

            public static void CreatelList(List<Order> orders, List<IUser> users)
            {
                // Створення об'єктів Excel
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlBook = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet xlSheet = (Excel.Worksheet)xlBook.Worksheets[1];

                // Додавання заголовків стовпців
                xlSheet.Cells[1, 1] = "Order ID";
                xlSheet.Cells[1, 2] = "User Name";
                xlSheet.Cells[1, 3] = "Type";
                xlSheet.Cells[1, 4] = "Date";
                xlSheet.Cells[1, 7] = "User ID";
                xlSheet.Cells[1, 8] = "User Name";
                xlSheet.Cells[1, 9] = "User Email";
                xlSheet.Cells[1, 10] = "User Phone";
                // Запис даних про замовлення у Excel
                for (int i = 0; i < orders.Count; i++)
                {
                    xlSheet.Cells[i + 2, 1] = orders[i].Id;
                    xlSheet.Cells[i + 2, 2] = orders[i].user.Name;
                    xlSheet.Cells[i + 2, 3] = orders[i].type.ToString();
                    xlSheet.Cells[i + 2, 4] = orders[i].Date.ToString();
                }
            // Запис даних про користувачів у Excel
            for (int i = 0; i < users.Count; i++)
            {
                xlSheet.Cells[i + 2, 7] = users[i].Id;
                xlSheet.Cells[i + 2, 8] = users[i].Name;
                xlSheet.Cells[i + 2, 9] = users[i].Email;
                xlSheet.Cells[i + 2, 10] = users[i].Phone;
            }
            // Збереження та закриття
            xlBook.SaveAs("orders.xlsx");
                xlBook.Close();
                xlApp.Quit();

                // Звільнення COM-об'єктів та очищення ресурсів
                Marshal.ReleaseComObject(xlSheet);
                Marshal.ReleaseComObject(xlBook);
                Marshal.ReleaseComObject(xlApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

     

}
