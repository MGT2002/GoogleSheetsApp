using System;
using System.Collections.Generic;
using System.Linq;
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
