using ClosedXML.Excel;
using ConsoleApp1;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PracticTask1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Введите ссылку на документ: ");
            string? path = Console.ReadLine();

            using (var workbook = new XLWorkbook(path))
            {
                var prod = workbook.Worksheet(1);
                var client = workbook.Worksheet(2);
                var order = workbook.Worksheet(3);

                StoreFile file = new StoreFile(client, prod, order) { };

                Console.Write("Введите наименование товара: ");
                string name = Console.ReadLine();

                file.Output(name);

                Console.Write("Введите наименование клиента: ");
                string company = Console.ReadLine();

                file.Update(company);
                workbook.SaveAs(path);

                Console.Write("Введите за какой период (месяц или год)  опредилить клиента: ");
                string inter = Console.ReadLine();

                order.Row(6).Hide();
                file.GoldenClient(inter);
              
                workbook.SaveAs(path);
            }
            Console.ReadKey();
        }
    }
}
