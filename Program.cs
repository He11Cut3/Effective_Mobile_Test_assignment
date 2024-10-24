using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using NLog;
using System.Globalization;

namespace DeliveryService
{
    class Program
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string logFilePath = string.Empty;
            string resultFilePath = string.Empty;
            string cityDistrict = string.Empty;
            DateTime? firstDeliveryDateTime = null;
            string dataFilePath = string.Empty;

            foreach (string arg in args)
            {
                if (arg.StartsWith("_cityDistrict="))
                {
                    cityDistrict = arg.Substring("_cityDistrict=".Length);
                }
                else if (arg.StartsWith("_firstDeliveryDateTime="))
                {
                    if (DateTime.TryParseExact(arg.Substring("_firstDeliveryDateTime=".Length), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDateTime))
                    {
                        firstDeliveryDateTime = parsedDateTime;
                    }
                }
                else if (arg.StartsWith("_deliveryLog="))
                {
                    logFilePath = arg.Substring("_deliveryLog=".Length);
                }
                else if (arg.StartsWith("_deliveryOrder="))
                {
                    resultFilePath = arg.Substring("_deliveryOrder=".Length);
                }
                else if (arg.StartsWith("_dataFile="))
                {
                    dataFilePath = arg.Substring("_dataFile=".Length);
                }
            }

            if (string.IsNullOrEmpty(logFilePath) || string.IsNullOrEmpty(resultFilePath) || string.IsNullOrEmpty(dataFilePath))
            {
                Console.WriteLine("Необходимо указать параметры _dataFile, _deliveryLog и _deliveryOrder.");
                Console.WriteLine("Пример: _dataFile=C:\\data.xlsx _deliveryLog=C:\\logs\\app.log _deliveryOrder=C:\\results\\orders.xlsx");
                return;
            }

            try
            {
                ConfigureLogger(logFilePath);
                logger.Info("Приложение запущено");

                if (string.IsNullOrEmpty(cityDistrict))
                {
                    Console.WriteLine("Введите район доставки:");
                    cityDistrict = Console.ReadLine();
                }

                if (firstDeliveryDateTime == null)
                {
                    Console.WriteLine("Введите начальное время доставки (в формате yyyy-MM-dd HH:mm:ss):");
                    if (!DateTime.TryParseExact(Console.ReadLine(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDateTime))
                    {
                        Console.WriteLine("Некорректный формат даты. Завершение работы.");
                        return;
                    }
                    firstDeliveryDateTime = parsedDateTime;
                }

                var orders = LoadOrders(dataFilePath);

                var filteredOrders = FilterOrders(orders, cityDistrict, firstDeliveryDateTime.Value);

                SaveToExcel(filteredOrders, resultFilePath);

                logger.Info("Обработка завершена успешно");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Произошла ошибка");
            }
        }

        static void ConfigureLogger(string logFilePath)
        {
            var config = new NLog.Config.LoggingConfiguration();
            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = logFilePath };
            config.AddRule(LogLevel.Info, LogLevel.Fatal, logfile);
            LogManager.Configuration = config;
        }

        static List<Order> LoadOrders(string filePath)
        {
            var orders = new List<Order>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    logger.Error("Файл Excel не содержит листов.");
                    return orders;
                }

                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text) ||
                        string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text) ||
                        string.IsNullOrWhiteSpace(worksheet.Cells[row, 3].Text) ||
                        string.IsNullOrWhiteSpace(worksheet.Cells[row, 4].Text))
                    {
                        logger.Warn($"Некорректные данные в строке {row}: одна из ячеек пуста.");
                        continue;
                    }

                    var order = new Order
                    {
                        OrderId = Guid.NewGuid(),
                        OrderNumber = worksheet.Cells[row, 1].Text,
                        District = worksheet.Cells[row, 3].Text
                    };

                    if (!double.TryParse(worksheet.Cells[row, 2].Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double weight))
                    {
                        logger.Warn($"Некорректный формат веса в строке {row}: {worksheet.Cells[row, 2].Text}");
                        continue;
                    }
                    order.Weight = weight;

                    string dateFormat = "yyyy-MM-dd HH:mm:ss";
                    if (!DateTime.TryParseExact(worksheet.Cells[row, 4].Text, dateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime deliveryTime))
                    {
                        logger.Warn($"Некорректный формат даты в строке {row}: {worksheet.Cells[row, 4].Text}");
                        continue;
                    }
                    order.DeliveryTime = deliveryTime;
                    orders.Add(order);
                }
            }

            return orders;
        }

        static List<Order> FilterOrders(List<Order> orders, string district, DateTime firstDeliveryTime)
        {
            var endTime = firstDeliveryTime.AddMinutes(30);
            return orders.Where(o => o.District.Equals(district, StringComparison.OrdinalIgnoreCase)
                                     && o.DeliveryTime >= firstDeliveryTime
                                     && o.DeliveryTime <= endTime)
                         .ToList();
        }

        static void SaveToExcel(List<Order> orders, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var worksheet = package.Workbook.Worksheets.Add("Заказы");
                worksheet.Cells[1, 1].Value = "Номер заказа";
                worksheet.Cells[1, 2].Value = "Вес";
                worksheet.Cells[1, 3].Value = "Район";
                worksheet.Cells[1, 4].Value = "Время доставки";

                for (int i = 0; i < orders.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = orders[i].OrderNumber;
                    worksheet.Cells[i + 2, 2].Value = orders[i].Weight;
                    worksheet.Cells[i + 2, 3].Value = orders[i].District;
                    worksheet.Cells[i + 2, 4].Value = orders[i].DeliveryTime;
                    worksheet.Cells[i + 2, 4].Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();


                package.SaveAs(new FileInfo(filePath));
                logger.Info($"Результаты сохранены в {filePath}");
            }
        }
    }

    public class Order
    {
        public Guid OrderId { get; set; }
        public string OrderNumber { get; set; }
        public double Weight { get; set; }
        public string District { get; set; }
        public DateTime DeliveryTime { get; set; }
    }

    public class Config
    {
        public string LogFilePath { get; set; }
        public string ResultFilePath { get; set; }
        public string CityDistrict { get; set; }
        public DateTime FirstDeliveryDateTime { get; set; }
    }
}
