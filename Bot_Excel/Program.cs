using System;
using System.ComponentModel;
using System.Data.Common;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using OfficeOpenXml;
using Telegram.Bot;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ConsoleApp2
{
    public class Program
    {
        private static readonly TelegramBotClient bot = new TelegramBotClient("7297731437:AAERIccwtDZnZnV3sNu2gjEpI5ze5Kq77uk");
        private static CancellationTokenSource cts = new CancellationTokenSource();
        private static readonly List<ProductInfo> productInfos = new List<ProductInfo>();

        static void Main(string[] args)
        {
            var receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = Array.Empty<UpdateType>() // получать все типы обновлений
            };

            bot.StartReceiving(HandleUpdateAsync, HandleErrorAsync, receiverOptions, cts.Token);
            Console.WriteLine("Bot is running...");
            Console.ReadLine();

            cts.Cancel();
        }
        private static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
        {
            try
            {
                var keyboard = new ReplyKeyboardMarkup(new[]
                                  {
                    new KeyboardButton("Экспорт")
                })
                {
                    ResizeKeyboard = true,
                    OneTimeKeyboard = false
                };
                if (update.Type == UpdateType.Message && update.Message!.Type == MessageType.Text)
                {


                    if (update.Message.Text == "/start")
                    {
                        await botClient.SendTextMessageAsync(update.Message.Chat.Id, "Привет! Отправьте данные, и нажмите кнопку 'Экспорт', чтобы экспортировать в Excel.", replyMarkup: keyboard, cancellationToken: cancellationToken);
                    }
                    else if (update.Message.Text == "Экспорт")
                    {
                        await botClient.SendTextMessageAsync(update.Message.Chat.Id, "Данные получены. Нажмите кнопку 'Экспорт', чтобы экспортировать в Excel.", cancellationToken: cancellationToken);
                        var filePath = GenerateExcelFile();
                        if (System.IO.File.Exists(filePath))
                        {
                            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                            {
                                var inputOnlineFile = InputFile.FromStream(stream, "data.xlsx");
                                await botClient.SendDocumentAsync(update.Message.Chat.Id, inputOnlineFile, cancellationToken: cancellationToken);
                            }
                        }
                        else
                        {
                            await botClient.SendTextMessageAsync(update.Message.Chat.Id, "Ошибка: файл не был создан.", cancellationToken: cancellationToken);
                        }

                        System.IO.File.Delete(filePath);
                        productInfos.Clear(); // очищаем список сообщений после экспорта
                    }
                    else
                    {
                        // Преобразование текста в объект
                        var productInfo = ParseTextToProductInfo(update.Message.Text);
                        if(productInfo.Package != null)
                        {
                            if (productInfo.Package.Length > 1 || productInfo.Package.Length == 1)
                            {
                                if (productInfo.Package.EndsWith("/"))
                                {
                                    productInfo.Package = productInfo.Package.Substring(0, productInfo.Package.Length - 1);
                                }
                            }
                        }
                        productInfos.Add(productInfo);
                    }

                }
            }
            catch (Exception ex)
            {
                await botClient.SendTextMessageAsync(update.Message.Chat.Id, $"Xato:{ex.Message} stacktrace:{ex.StackTrace}", cancellationToken: cancellationToken);

                Console.WriteLine($"An error occurred_catch: {ex.Message}");
                Console.WriteLine($"\nStackTrace_catch: {ex.StackTrace}");
                Console.WriteLine($"\nSource_catch: {ex.Source}");
            }
            
        }
        private static Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
        {
            Console.WriteLine($"\nAn error occurred: {exception.Message}");
            Console.WriteLine($"\nStackTrace: {exception.StackTrace}");
            Console.WriteLine($"\nSource: {exception.Source}");
            //cts.Cancel();
            return Task.CompletedTask;
        }

        private static string GenerateExcelFile()
        {
            // Устанавливаем лицензионный контекст для EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var filePath = Path.Combine(Path.GetTempPath(), "data.xlsx");

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells[1, 1].Value = "№";
                worksheet.Cells[1, 2].Value = "Махсулот";
                worksheet.Cells[1, 3].Value = "Мутахасис";
                worksheet.Cells[1, 4].Value = "Мижоз";
                worksheet.Cells[1, 5].Value = "Номер";
                worksheet.Cells[1, 6].Value = "Манзил";
                worksheet.Cells[1, 7].Value = "Вилоят";
                worksheet.Cells[1, 8].Value = "Логитсика";
                worksheet.Cells[1, 9].Value = "Нархи";
                worksheet.Cells[1, 10].Value = "УП";
                worksheet.Cells[1, 11].Value = "Источник";

                // Заполнение данными
                for (int i = 0; i < productInfos.Count; i++)
                {
                    var productInfo = productInfos[i];
                    worksheet.Cells[i + 2, 1].Value = productInfo.Order;
                    worksheet.Cells[i + 2, 2].Value = productInfo.Product;
                    worksheet.Cells[i + 2, 3].Value = productInfo.Specialist;
                    worksheet.Cells[i + 2, 4].Value = productInfo.Client;
                    if(productInfo.Numbers.Count > 1 && productInfo.Numbers != null)
                    {
                        worksheet.Cells[i + 2, 5].Value = string.Join("  ", productInfo.Numbers);
                    }
                    else
                    {
                        worksheet.Cells[i + 2, 5].Value = string.Join(" ", productInfo.Numbers);
                    }
                    worksheet.Cells[i + 2, 6].Value = productInfo.City;
                    worksheet.Cells[i + 2, 7].Value = productInfo.Adress;
                    worksheet.Cells[i + 2, 8].Value = productInfo.Logistics;
                    worksheet.Cells[i + 2, 9].Value = productInfo.Price;
                    worksheet.Cells[i + 2, 10].Value = productInfo.Package;
                    worksheet.Cells[i + 2, 11].Value = productInfo.Source;
                }
                package.Save();
            }
            return filePath;
        }
        private static ProductInfo ParseTextToProductInfo(string text)
        {
            var lines = text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var productInfo = new ProductInfo();
            string currentKey = null;

            foreach (var line in lines)
            {
                if (line.Contains(":"))
                {
                    var parts = line.Split(new[] { ':' }, 2);
                    currentKey = parts[0].Trim();
                    var value = parts[1].Trim();

                    switch (currentKey.ToUpper())
                    {
                        case "МАХСУЛОТ":
                            productInfo.Product = value;
                            break;
                        case "МУТАХАСИС":
                            productInfo.Specialist = value;
                            break;
                        case "МИЖОЗ":
                            productInfo.Client = value;
                            break;
                        case "НОМЕР":
                            productInfo.Numbers.Add(value);
                            break;
                        case "МАНЗИЛ":
                            productInfo.City = value;
                            break;
                        case "НАРХИ":
                            productInfo.Package = value;
                            break;
                        case "ЛОГИСТИКА":
                            productInfo.Logistics = value;
                            break;
                        default:
                            productInfo.Source += line + " ";
                            break;
                    }
                }
                else
                {
                    if (currentKey != null)
                    {
                        switch (currentKey.ToUpper())
                        {
                            case "НОМЕР":
                                productInfo.Numbers.Add(line.Trim());
                                break;
                            case "МАНЗИЛ":
                                productInfo.Adress += " " + line.Trim();
                                break;
                            case "НАРХИ":
                                productInfo.Price += " " + line.Trim();
                                break;
                            default:
                                productInfo.Source += line + " ";
                                break;
                        }
                    }
                    else
                    {
                        productInfo.Order += line + " ";
                    }
                }

            }
            return productInfo;
        }
    }
    public class ProductInfo
    {
        public string? Order { get; set; }
        public string? Product { get; set; }
        public string? Specialist { get; set; }
        public string? Client { get; set; }
        public List<string?> Numbers { get; set; } = new List<string?>();
        public string? City { get; set; }
        public string? Adress { get; set; }
        public string? Package { get; set; }
        public string? Price { get; set; }
        public string? Logistics { get; set; }
        public string? Source { get; set; }
    }
}

