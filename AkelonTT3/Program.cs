using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using ClosedXML.Excel;

namespace AkelonTT3
{
    class Program
    {
        static void Main(string[] args)
        {
            WorkCycle();
        }

        static void WorkCycle()
        {
            Console.WriteLine("Введите путь до рабочего файла:");
            string? filePath = Console.ReadLine();
            using (var book = new XLWorkbook(filePath))
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
                                    GetClientRequestsData(book, productName);
                                    case1 = false;
                                }
                            }
                            case1 = true;
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
                                        ChangeClientContact(book, client, contact);
                                        case2 = false;
                                    }
                                }
                            }
                            case2 = true;
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
                                        GetGoldenClient(book, yearNum, monthNum);
                                        case3 = false;
                                    }
                                }
                            }
                            case3 = false;
                            break;
                        case "4":
                            on = false;
                            break;
                        default:
                            Console.WriteLine("Неизвестная команда.");
                            break;
                    }
                }
            }
        }

        static void GetClientRequestsData(XLWorkbook wb, string productName)
        {
            // Получаем таблицу товары
            var wsProducts = wb.Worksheet("Товары");
            // вытаскиваем все заполненные строки кроме шапки
            var rowsP = wsProducts.RangeUsed().RowsUsed().Skip(1);
            // получаем строку с нужным товаром
            var targetProductRow = rowsP.FirstOrDefault(r => r.Cell(2).GetValue<string>() == productName, null);
            if (targetProductRow == null)
            {
                Console.WriteLine("Товар не найден");
                return;
            }
            // получаем код товаром
            int productCode = targetProductRow.Cell(1).GetValue<int>();
            // получаем таблицу заявки
            var wsRequests = wb.Worksheet("Заявки");
            // таблицу без шапки
            var rowsR = wsRequests.RangeUsed().RowsUsed().Skip(1);
            // строки с кодом товара
            var requestRows = rowsR.Where(r => r.Cell(2).GetValue<int>() == productCode).ToList();
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
            // получаем данные о клиентах
            var wsClients = wb.Worksheet("Клиенты");
            // таблицу без шапки
            var rowsC = wsClients.RangeUsed().RowsUsed().Skip(1);
            //строки с кодом клиента
            List<IXLRangeRow> clients = new List<IXLRangeRow>();
            IXLRangeRow? clRow = null;
            for(int i = 0; i < clientsId.Count; i++)
            {
                clRow = rowsC.FirstOrDefault(r => r.Cell(1).GetValue<int>() == clientsId[i], null);
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
                Console.WriteLine($"Стоимость заказа: {targetProductRow.Cell(4).GetValue<decimal>() *
                    (decimal)requestRows[j].Cell(5).GetValue<int>()} [Рублей]");
                j++;
            }
        }

        static void ChangeClientContact(XLWorkbook wb, string client, string contact)
        {
            // получаем данные о клиентах
            var wsClients = wb.Worksheet("Клиенты");
            var rows = wsClients.RangeUsed().RowsUsed().Skip(1);
            var targetClientRow = rows.FirstOrDefault(r => r.Cell(2).GetValue<string>() == client, null);
            if (targetClientRow != null) 
            {
                string oldContact = targetClientRow.Cell(4).GetValue<string>();
                targetClientRow.Cell(4).SetValue(contact);
                wb.Save();
                Console.WriteLine($"Было изменено контактное лицо организации {client}.");
                Console.WriteLine($"Старое контактное лицо: {oldContact}");
                Console.WriteLine($"Новое контактное лицо: {targetClientRow.Cell(4).GetValue<string>()}");
            }
            else
            {
                Console.WriteLine($"Клиент {client} не найден");
            }
        }

        static void GetGoldenClient(XLWorkbook wb, int year, int month)
        {
            // таблица с заказами
            var wsRequests = wb.Worksheet("Заявки");
            var rowsR = wsRequests.RangeUsed().RowsUsed().Skip(1);
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
            var checkedRequests = rowsR.Where(r => r.Cell(6).GetValue<DateTime>() >= dtStart &&
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
            var wsClients = wb.Worksheet("Клиенты");
            var goldenClient = wsClients.RangeUsed().RowsUsed().Skip(1)
                .First(r => r.Cell(1).GetValue<int>() == result.Key);
            Console.WriteLine("Золтой клиент:");
            Console.WriteLine($"Организация {goldenClient.Cell(2).GetValue<string>()}");
        }
    }
}
