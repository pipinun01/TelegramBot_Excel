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
        private static readonly TelegramBotClient bot = new TelegramBotClient("6403014265:AAHDsNkXlkSR4xFB07HL9uc2Yr1voTX-pHc");
        private static CancellationTokenSource cts = new CancellationTokenSource();
        private static readonly List<string> Messages = new List<string>();

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

                        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            var inputOnlineFile = InputFile.FromStream(stream, "data.xlsx");
                            await botClient.SendDocumentAsync(update.Message.Chat.Id, inputOnlineFile, cancellationToken: cancellationToken);
                        }

                        System.IO.File.Delete(filePath);
                        Messages.Clear(); // очищаем список сообщений после экспорта
                    }
                    else
                    {
                        Messages.Add(update.Message.Text);
                        var data = new Dictionary<string, string>();
                        var lines = update.Message.Text.Split('\n');
                        //foreach (var line in lines)
                        //{
                        //    var parts = line.Split(':');
                        //    if (parts.Length == 2)
                        //    {
                        //        data[parts[0].Trim()] = parts[1].Trim();
                        //    }
                        //}
                        //var jsonSerialize = JsonConvert.SerializeObject(data, Formatting.Indented);
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

                int row = 2;
                foreach (var message in Messages)
                {
                    var data = ParseMessage(message);
                    worksheet.Cells[row, 1].Value = data["Ракам"];
                    worksheet.Cells[row, 2].Value = data["Махсулот"];
                    worksheet.Cells[row, 3].Value = data["Мутахасис"];
                    worksheet.Cells[row, 4].Value = data["Мижоз"];
                    worksheet.Cells[row, 5].Value = data["Номер"];
                    worksheet.Cells[row, 6].Value = data["Манзил"];
                    worksheet.Cells[row, 7].Value = data["Вилоят"];
                    worksheet.Cells[row, 8].Value = data["Логистика"];
                    worksheet.Cells[row, 9].Value = data["Нархи"];
                    worksheet.Cells[row, 10].Value = data["Упаковка"];
                    worksheet.Cells[row, 11].Value = data["Источник"];
                    row++;
                }

                package.Save();
            }
            return filePath;
        }
        private static Dictionary<string, string> ParseMessage(string message)
        {
            var data = new Dictionary<string, string>
            {
                ["Ракам"] = ExtractData(message, @"^(\d+)"),
                ["Махсулот"] = ExtractData(message, @"Махсулот:\s*(.*)"),
                ["Мутахасис"] = ExtractData(message, @"Мутахасис\s*:\s*(.*)"),
                ["Мижоз"] = ExtractData(message, @"Мижоз\s*:\s*(.*)"),
                ["Номер"] = ExtractData(message, @"Номер\s*:\s*([\d\s]+)").Replace("\n", "  "),
                //["Манзил"] = ExtractData(message, @"Манзил\s*:\s*(.*)").Replace("\n", " "),
                ["Упаковка"] = ExtractData(message, @"Нархи:\s*(\d+)/").Trim(),
                ["Логистика"] = ExtractData(message, @"Логистика\s*:\s*([\d\s]+)").Replace("\n", "  "),
                ["Нархи"] = ExtractData(message, @"Нархи:\s*\d+/\s*([^\r\n]+)"),
                ["Источник"] = ExtractData(message, @"#\s*(.*)")
            };
            var address = ExtractData(message, @"Манзил:\s*([^\r\n]+)[\r\n\s]*([^\r\n]*)");
            var addressParts = address.Split(new[] { '\n', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (addressParts.Length >= 2)
            {
                data["Манзил"] = addressParts[1].Trim();
                data["Вилоят"] = addressParts[0].Trim();
            }
            else
            {
                if (addressParts.Length == 1)
                {
                    data["Манзил"] = string.Empty;
                    data["Вилоят"] = addressParts[0].Trim();
                }
                else if(addressParts.Length == 0)
                {
                    data["Манзил"] = string.Empty;
                    data["Вилоят"] = "";
                }
                else
                {

                    data["Манзил"] = string.Empty;
                    data["Вилоят"] = addressParts[0].Trim();
                }
                
            }

            return data;
        }

        private static string ExtractData(string text, string pattern)
        {
            var match = Regex.Match(text, pattern, RegexOptions.Multiline |RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
        }
    }
}

