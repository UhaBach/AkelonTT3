using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Math;

namespace AkelonTT3
{
    class Program
    {
        static XMLFileReader reader;
        static XMLFileWriter writer;

        static void Main(string[] args)
        {
            bool on = true;
            while (on)
            {
                Console.WriteLine("Введите путь до рабочего файла:");
                string? filePath = Console.ReadLine();
                if (filePath != null)
                {
                    try
                    {
                        var wb = new XLWorkbook(filePath);
                        reader = new XMLFileReader(wb);
                        writer = new XMLFileWriter(wb);
                        on = false;
                    }
                    catch
                    {
                        Console.WriteLine("Неверный путь до файла или файл уже используется другим процессом.");
                        Console.WriteLine("Попробуйте снова.");
                        Console.WriteLine("==================================================================");
                        break;
                    }
                }
            }
            WorkCycle();
        }

        static void WorkCycle()
        {
            bool on = true, case1 = true, case2 = true, case3 = true;
            string? option, productName, client, contact, year, month;
            while (on)
            {
                Console.WriteLine("Что вы хотите сделать? (введите номер команды)");
                Console.WriteLine("1 - по наименованию товара получить данные о клиентах, заказавших товар.");
                Console.WriteLine("2 - запрос на изменение контактного лица клиента.");
                Console.WriteLine("3 - определить золотого клиента.");
                Console.WriteLine("4 - завершить работу приложения.");
                Console.WriteLine("==================================================================");
                option = Console.ReadLine();
                switch (option)
                {
                    case "1":
                        while (case1)
                        {
                            Console.WriteLine("Введите название товара:");
                            productName = Console.ReadLine();
                            if (productName != null)
                            {
                                GetClientRequestsData(productName);
                                case1 = false;
                            }
                        }
                        case1 = true;
                        productName = null;
                        Console.WriteLine("==================================================================");
                        break;
                    case "2":
                        while (case2)
                        {
                            Console.WriteLine("Введите название организации:");
                            client = Console.ReadLine();
                            if (client != null)
                            {
                                Console.WriteLine("Введите ФИО контактного лица организации:");
                                contact = Console.ReadLine();
                                if (contact != null)
                                {
                                    ChangeClientContact(client, contact);
                                    case2 = false;
                                }
                            }
                        }
                        case2 = true;
                        client = null;
                        contact = null;
                        Console.WriteLine("==================================================================");
                        break;
                    case "3":
                        while (case3)
                        {
                            Console.WriteLine("Введите год, по которому производится выборка:");
                            year = Console.ReadLine();
                            if (int.TryParse(year, out int yearNum))
                            {
                                Console.WriteLine("Введите мемсяц, по которому производится выборка, в числовом виде:");
                                month = Console.ReadLine();
                                if (int.TryParse(month, out int monthNum))
                                {
                                    GetGoldenClient(yearNum, monthNum);
                                    case3 = false;
                                }
                            }
                        }
                        case3 = false;
                        year = null;
                        month = null;
                        Console.WriteLine("==================================================================");
                        break;
                    case "4":
                        on = false;
                        break;
                    default:
                        Console.WriteLine("Неизвестная команда.");
                        Console.WriteLine("==================================================================");
                        break;
                }
            }
        }

        static void GetClientRequestsData(string productName)
        {
            // получаем таблицу товары без шапки
            var rowsProduct = reader.GetSheetRows("Товары");
            var targetProductRow = rowsProduct.FirstOrDefault(r => r.Cell(2).GetValue<string>() == productName, null);
            if (targetProductRow == null)
            {
                Console.WriteLine("Товар не найден");
                return;
            }
            int productCode = targetProductRow.Cell(1).GetValue<int>();
            // получаем таблицу заявки без шапки
            var rowsRequest = reader.GetSheetRows("Заявки");
            var requestRows = rowsRequest.Where(r => r.Cell(2).GetValue<int>() == productCode).ToList();
            if (requestRows.Count == 0)
            {
                Console.WriteLine("Заявок по данному товару не найдено");
                return;
            }
            // сохраняем все коды клиентов
            List<int> clientsId = new List<int>();
            int id = 0;
            for(int i = 0; i < requestRows.Count; i++)
            {
                id = requestRows[i].Cell(3).GetValue<int>();
                if (!clientsId.Contains(id)) clientsId.Add(id);
            }
            // получаем таблицу клиенты без шапки
            var rowsClient = reader.GetSheetRows("Клиенты");
            List<IXLRangeRow> clients = new List<IXLRangeRow>();
            IXLRangeRow? clRow = null;
            for(int i = 0; i < clientsId.Count; i++)
            {
                clRow = rowsClient.FirstOrDefault(r => r.Cell(1).GetValue<int>() == clientsId[i], null);
                if (clRow != null) clients.Add(clRow);
            }
            Console.WriteLine($"Товар {productName}:");
            int j = 0;
            foreach (var client in clients)
            {
                Console.WriteLine($"\tЗаказчик: {client.Cell(2).GetValue<string>()}, {client.Cell(3).GetValue<string>()}");
                Console.WriteLine($"\tКонтактное лицо: {client.Cell(4).GetValue<string>()}");
                Console.WriteLine($"\tДата заказа: {requestRows[j].Cell(6).GetValue<DateTime>().ToString()}");
                Console.WriteLine($"\tОбъём заказа: {requestRows[j].Cell(5).GetValue<int>()} " +
                    $"[{targetProductRow.Cell(3).GetValue<string>()}]");
                Console.WriteLine($"\tСтоимость заказа: {targetProductRow.Cell(4).GetValue<decimal>() *
                    (decimal)requestRows[j].Cell(5).GetValue<int>()} [Рублей]");
                j++;
            }
        }

        static void ChangeClientContact(string client, string contact)
        {
            // получаем таблицу клиенты без шапки
            var rowsClient = reader.GetSheetRows("Клиенты");
            var targetClientRow = rowsClient.FirstOrDefault(r => r.Cell(2).GetValue<string>() == client, null);
            if (targetClientRow != null) 
            {
                string oldContact = targetClientRow.Cell(4).GetValue<string>();
                writer.WtiteDataInCell(targetClientRow.Cell(4), contact);
                writer.SaveFileData();
                Console.WriteLine($"Было изменено контактное лицо организации {client}.");
                Console.WriteLine($"\tСтарое контактное лицо: {oldContact}");
                Console.WriteLine($"\tНовое контактное лицо: {targetClientRow.Cell(4).GetValue<string>()}");
            }
            else
            {
                Console.WriteLine($"Клиент {client} не найден");
            }
        }

        static void GetGoldenClient(int year, int month)
        {
            // таблица с заказами без шапки
            var rowsRequest = reader.GetSheetRows("Заявки");
            DateTime dtStart = new DateTime(year, month, 1);
            if (month == 12)
            {
                month = 1;
                year += 1;
            }
            else
            {
                month++;
            }
            DateTime dtEnd = new DateTime(year, month, 1);
            var checkedRequests = rowsRequest.Where(r => r.Cell(6).GetValue<DateTime>() >= dtStart &&
                r.Cell(6).GetValue<DateTime>() < dtEnd).ToList();
            // сохраняем уникальные id клиентов в список
            List<int> clientsId = new List<int>();
            foreach (var req in checkedRequests)
            {
                clientsId.Add(req.Cell(3).GetValue<int>());
            }
            clientsId = clientsId.Distinct().ToList();
            // сохраняем кол-во повторений для каждого id клиета в словарь
            Dictionary<int, int> dict = new Dictionary<int, int>();
            foreach (var id in clientsId)
            {
                dict[id] = 0;
            }
            foreach (var req in checkedRequests)
            {
                dict[req.Cell(3).GetValue<int>()] += 1;
            }
            // в задании не было уточнено может ли быть более одного золотого клиента
            // я решил взять первого попавшегося
            var result = dict.OrderByDescending(r => r.Value).First();
            // вытаскиваем злотого клиента из таблицы с клиентами
            var goldenClient = reader.GetSheetRows("Клиенты")
                .First(r => r.Cell(1).GetValue<int>() == result.Key);
            Console.WriteLine("Золтой клиент:");
            Console.WriteLine($"\tОрганизация {goldenClient.Cell(2).GetValue<string>()}");
        }
    }
}
