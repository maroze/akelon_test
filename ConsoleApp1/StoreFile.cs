using ClosedXML.Excel;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ConsoleApp1
{
    public class StoreFile
    {
        private IXLWorksheet wsClient;
        private IXLWorksheet wsProduct;
        private IXLWorksheet wsOrder;
        public StoreFile(IXLWorksheet _wsClient, IXLWorksheet _wsProduct, IXLWorksheet _wsOrder) 
        { 
            this.wsClient = _wsClient;
            this.wsProduct = _wsProduct;
            this.wsOrder = _wsOrder;
        }
        public void Output(string name)
        {
            while (name != "0")
            {
                int idProduct;
                int idOrder;
                int idClient;
                for (int i = 1; i <= wsProduct.RowsUsed().Count(); i++)
                {
                    if (wsProduct.Cell(i, 2).Value.Equals(name))
                    {
                        idProduct = (int)wsProduct.Cell(i, 1).Value;

                        for (int j = 1; j <= wsOrder.RowsUsed().Count(); j++)
                        {
                            if (wsOrder.Cell(j, 2).Value.Equals(idProduct))
                            {
                                idClient = (int)wsClient.Cell(j, 1).Value;
                                idOrder = (int)wsOrder.Cell(j, 1).Value;
                                Console.WriteLine($"Информация по товару: {wsClient.Row(idClient).Cell(2).GetString()}, {wsClient.Cell(idClient, 3).Value}," +
                                    $"{wsClient.Cell(idClient, 4).Value}, {wsOrder.Cell(idOrder, 5).Value}, " +
                                    $"{wsOrder.Cell(idOrder, 6).Value}");
                            }
                        }
                    }
                }
                Console.Write("Для завершения введите 0: ");
                name = Console.ReadLine();
            }
        }

        public void Update(string company)
        {
            for (int i = 1; i <= wsClient.RowsUsed().Count(); i++)
            {
                if (wsClient.Cell(i, 2).Value.Equals(company))
                {
                    Console.Write("Введите контактное лицо клиента: ");
                    string contact = Console.ReadLine();

                    wsClient.Cell(i, 4).Value = contact;
                    Console.WriteLine($"Контактное лицо {wsClient.Cell(i, 4).Value} организации {wsClient.Cell(i, 2).Value}"); 
                }
            }
        }
        public void GoldenClient2()
        {
            wsOrder.AutoFilter.Sort(1);

            Dictionary<string, int> Count = new Dictionary<string, int>();
            for (int i = 1; i < wsOrder.RowsUsed().Count(); i++)
            {
                for (int j = i; j < wsOrder.RowsUsed().Count(); j++)
                {
                    if (!wsOrder.Cell(i, 3).Value.Equals(wsOrder.Cell(j, 3).Value))
                    {
                        Count.Add(wsOrder.Cell(i, 3).Value.ToString(), j - i);
                        i = j - 1;
                        break;
                    }
                    if (j == wsOrder.RowsUsed().Count() - 1)
                    {
                        Count.Add(wsOrder.Cell(i, 3).Value.ToString(), j - i + 1);
                        i = j;
                    }
                }
            }   
            wsOrder.Columns().AdjustToContents();
            wsOrder.Row(6).Hide();
            //не дописан

        }
        public void GoldenClient(string inter)
        {
            DateTime date;
            switch (inter)
            {
                case "месяц":
                    Console.Write("Введите месяц и год в формате мм.гггг ");
                    int[] mdt = Array.ConvertAll(Console.ReadLine().Split('.'), int.Parse);
                    date = new DateTime(mdt[1], mdt[0], 01);
                    wsOrder.RangeUsed().SetAutoFilter().Column(6).AddDateGroupFilter(date, XLDateTimeGrouping.Month);
                    GoldenClient2();
                    break;
                case "год":
                    Console.Write("Введите год в формате гггг ");
                    int[] ydt = Array.ConvertAll(Console.ReadLine().Split(), int.Parse);
                    date = new DateTime(ydt[0], 01, 01);
                    wsOrder.RangeUsed().SetAutoFilter().Column(6).AddDateGroupFilter(date, XLDateTimeGrouping.Year);
                    GoldenClient2();
                    break;
            }
        }
    }
}
