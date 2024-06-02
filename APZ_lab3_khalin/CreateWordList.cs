using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using Khalin_Kypcova_612pst.Classes;

namespace APZ_lab3_khalin
{
    public static class CreateWordList
    {

            public static void CreateList(List<Order> orders, List<IUser> users)
            {
                // Створення об'єктів Word
                Word.Application wdApp = new Word.Application();
                Word.Document doc = wdApp.Documents.Add();
                Word.Paragraph para;

                // Додавання заголовка
                para = doc.Content.Paragraphs.Add();
                para.Range.Text = "Order List";
                para.Range.Font.Bold = 1;
                para.Range.Font.Size = 24;
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.InsertParagraphAfter();


                // Додавання даних про замовлення у Word
                foreach (var order in orders)
                {
                    para = doc.Content.Paragraphs.Add();
                    para.Range.Text = $"Order ID: {order.Id}";
                    para.Range.Font.Bold = 0;
                    para.Range.Font.Size = 12;
                    para.Range.InsertParagraphAfter();

                    para = doc.Content.Paragraphs.Add();
                    para.Range.Text = $"User Name: {order.user.Name}";
                    para.Range.InsertParagraphAfter();

                    para = doc.Content.Paragraphs.Add();
                    para.Range.Text = $"Type: {order.type}";
                    para.Range.InsertParagraphAfter();

                    para = doc.Content.Paragraphs.Add();
                    para.Range.Text = $"Date: {order.Date}";
                    para.Range.InsertParagraphAfter();

                    para.Range.InsertParagraphAfter();
                }
                // Додавання заголовка
                para = doc.Content.Paragraphs.Add();
                para.Range.Text = "User List";
                para.Range.Font.Bold = 1;
                para.Range.Font.Size = 24;
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.InsertParagraphAfter();
                foreach (var user in users)
                {
                    para = doc.Content.Paragraphs.Add();
                    para.Range.Text = $"User ID: {user.Id}";
                    para.Range.Font.Bold = 0;
                    para.Range.Font.Size = 12;
                    para.Range.InsertParagraphAfter();

                    para = doc.Content.Paragraphs.Add();
                    para.Range.Text = $"User Name: {user.Name}";
                    para.Range.InsertParagraphAfter();

                    para = doc.Content.Paragraphs.Add();
                    para.Range.Text = $"User Email: {user.Email}";
                    para.Range.InsertParagraphAfter();

                    para = doc.Content.Paragraphs.Add();
                    para.Range.Text = $"User Phone: {user.Phone}";
                    para.Range.InsertParagraphAfter();

                    para.Range.InsertParagraphAfter();
                }

                // Збереження та закриття документа
                object filename = "orders.docx";
                doc.SaveAs2(ref filename);
                doc.Close();
                wdApp.Quit();

                // Звільнення COM-об'єктів та очищення ресурсів
                Marshal.ReleaseComObject(para);
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wdApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


}
