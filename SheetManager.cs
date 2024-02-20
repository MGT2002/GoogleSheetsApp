using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using static Google.Apis.Sheets.v4.SpreadsheetsResource
    .ValuesResource.GetRequest;

namespace GoogleSheetsApp
{
    internal class SheetManager
    {
        private const string apiKey = "AIzaSyBUU9R7_wicSiNDiR1kTF4ihiJkC7Hssf4";
        private readonly static SheetsService service;
        public string SheetUrl { get; }
        public string SpreadsheetId { get; }

        static SheetManager()
        {
            // Create a Sheets service object
            service = new SheetsService(new BaseClientService.Initializer()
            {
                ApplicationName = "GoogleSheetsApp",
                ApiKey = apiKey
            });

        }
        public SheetManager(string sheetUrl)
        {
            SheetUrl = sheetUrl;
            SpreadsheetId = sheetUrl.Split('/')[5];
        }

        /// <summary>Gets data from your SheetUrl</summary>
        /// <param name="range">Sheets API range e.g. -> "A1:ZZ"</param>
        public ValueRange? GetData(string range, bool getFormula = false)
        {
            // Read data from the spreadsheet
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);
            if (getFormula)
            {
                request.ValueRenderOption = ValueRenderOptionEnum.FORMULA;
            }
            ValueRange response;

            try
            {
                response = request.Execute();
            }
            catch (AggregateException e)
            {
                foreach (var i in e.InnerExceptions)
                    Console.WriteLine("ERROR: " + i.Message);
                return null;
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: " + e.Message);
                return null;
            }

            return response;
        }

        /// <param name="userEnteredFormat">Set false to get effective format</param>
        /// <returns>Array that has format of each column in its elements</returns>
        public string[]? GetColFormats(bool userEnteredFormat = true)
        {
            string[]? colFormats = default;

            var request = service.Spreadsheets.Get(SpreadsheetId);
            request.IncludeGridData = true;

            try
            {
                var response = request.Execute();

                foreach (var sheet in response.Sheets)
                {
                    RowData row = sheet.Data[0].RowData[1];//B1:B row
                    colFormats = new string[row.Values.Count];
                    for (int i = 0; i < colFormats.Length; i++)
                    {
                        CellFormat format = userEnteredFormat ?
                            row.Values[i].UserEnteredFormat :
                            row.Values[i].EffectiveFormat;
                        colFormats[i] = format.NumberFormat?.Type ?? "TEXT";
                    }
                }

                return colFormats;
            }
            catch (AggregateException e)
            {
                foreach (var i in e.InnerExceptions)
                    Console.WriteLine("ERROR: " + i.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: " + e.Message);
            }

            return default;
        }
        public static bool IsGoogleSheetsUrl(string s)
        {
            if (s == "")
                return false;

            if (!s.StartsWith("https://docs.google.com/spreadsheets/d/"))
                return false;

            // Extract potential ID part
            var parts = s.Split('/');
            if (parts.Length < 6)
                return false;

            return true;
        }
    }
}
