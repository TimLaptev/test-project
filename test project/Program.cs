using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Types;
using Telegram.Bot.Types.ReplyMarkups;
using Telegram.Bot.Args;
using Telegram.Bot.Polling;

namespace SheetsQuickstart
{
    class Program
    {
        private static string tiken = "6147084307:AAEcI6NRi7ZT7z3Fa9DD4sTaIYpqOAyP4iY";
        private static TelegramBotClient client_bot;
        
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "p";

        static void Main(string[] args)
        {

             client_bot = new TelegramBotClient(tiken);
            var cts = new CancellationTokenSource();
            var cancellationToken = cts.Token;
            var receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = { }, // receive all update types
            };

            client_bot.StartReceiving(HandleUpdateAsync,
                HandleErrorAsync,
                receiverOptions,
                cancellationToken

                );
            
            Console.Read();

           

            
            



            }

        public static async Task HandleErrorAsync(ITelegramBotClient arg1, Exception arg2, CancellationToken arg3)
        {
            Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(arg2));

        }

        private static async Task HandleUpdateAsync(ITelegramBotClient arg1, Update arg2, CancellationToken arg3)
        {

            UserCredential credential;

            using (var stream = new FileStream(@"C:\Users\Тимуратор\source\repos\test project\test project\client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                        System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Создать сервис Google Sheets API.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Определить параметры запроса.
            String spreadsheetId = "18QMBfkp0Ihmm0PNoUO4v5er0Tj8lVu8JHnOiTtOTPH0";
            String range = "2023!A:H";

            SpreadsheetsResource.ValuesResource.GetRequest request =
            service.Spreadsheets.Values.Get(spreadsheetId, range);



            ValueRange response = request.Execute();
            IList<IList<Object>> values_1 = response.Values;

            request = service.Spreadsheets.Values.Get(spreadsheetId, range);










            DateTime sys = DateTime.Now;
            List<string> client = new List<string>();
            List<string> client_today = new List<string>();
            List<string> client_tomorrow = new List<string>();
            List<string> client_after_tomorrow = new List<string>();

            if (values_1 != null && values_1.Count > 0)
            {

                for (int i = 0; i < values_1.Count; i++)
                {

                }
                foreach (var row in values_1)
                {
                    try
                    {
                        if (row.ToString() != "")
                        {
                            row.Remove(0);
                            row.Remove(7);
                        }
                        string a = (string)row[0];
                        string h = (string)row[7];

                        DateTime dH = DateTime.Parse(h);
                        if (sys.ToString("dd.MM.yyyy") == h)
                        {
                            client_today.Add(a + "  " + h);

                        }
                        DateTime t_sys = sys.Date.AddDays(1);

                        if (t_sys.ToString("dd.MM.yyyy") == h)
                        {
                            client_tomorrow.Add(a + "  " + h);
                        }

                        DateTime tt_sys = sys.Date.AddDays(2);


                        if (tt_sys.ToString("dd.MM.yyyy") == h)
                        {
                            client_after_tomorrow.Add(a + "  " + h);
                        }

                        if (dH > tt_sys)
                        {
                            client.Add(a + "  " + h);
                        }

                    }

                    catch (Exception)
                    {

                    }


                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }





           

            var msg = arg2.Message;
            if (msg.Text != null)
            {

                switch (msg.Text)
                {
                    case "/start":
                        {

                            await client_bot.SendTextMessageAsync(msg.Chat.Id, "Рабочий бот Belebey LAB", replyMarkup: GetButtons());



                            break;
                        }
                    case "Сегодня":
                        {
                            if (client_today.Count == 0)
                            {
                                await client_bot.SendTextMessageAsync(msg.Chat.Id, "нет работ", replyMarkup: GetButtons());
                            }

                            else
                            {
                                foreach (var item in client_today)
                                {
                                    await client_bot.SendTextMessageAsync(msg.Chat.Id, item, replyMarkup: GetButtons());
                                }
                            }

                            break;
                        }
                    case "Завтра":
                        {
                            if (client_tomorrow.Count == 0)
                            {
                                await client_bot.SendTextMessageAsync(msg.Chat.Id, "нет работ", replyMarkup: GetButtons());
                            }
                            else
                            {

                                foreach (var item in client_tomorrow)
                                {
                                    await client_bot.SendTextMessageAsync(msg.Chat.Id, item, replyMarkup: GetButtons());
                                }
                            }
                            break;
                        }

                    case "После завтра":
                        {
                            if (client_after_tomorrow.Count == 0) { await client_bot.SendTextMessageAsync(msg.Chat.Id, "нет работ", replyMarkup: GetButtons()); }
                            else
                            {
                                foreach (var item in client_after_tomorrow)
                                {
                                    await client_bot.SendTextMessageAsync(msg.Chat.Id, item, replyMarkup: GetButtons());
                                }
                            }
                            break;
                        }
                    case "Остальные дни":
                        {
                            if (client.Count == 0)
                            {
                                await client_bot.SendTextMessageAsync(msg.Chat.Id, "нет работ", replyMarkup: GetButtons());
                            }
                            else
                            {
                                foreach (var item in client)
                                {
                                    await client_bot.SendTextMessageAsync(msg.Chat.Id, item, replyMarkup: GetButtons());
                                }
                            }

                            break;
                        }
                }


            }


        }

       

        private static IReplyMarkup GetButtons()
        {
            return new ReplyKeyboardMarkup(tiken)
            {
                Keyboard = new List<List<KeyboardButton>>
                {
                    new List<KeyboardButton>{ new KeyboardButton("Сегодня"),new KeyboardButton("Завтра") },
                     new List<KeyboardButton>{ new KeyboardButton ("После завтра"),new KeyboardButton ("Остальные дни") },

                }
            };
        }
    }
    
}