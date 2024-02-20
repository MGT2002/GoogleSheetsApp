using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheetsApp
{
    internal static class FormatToSqlConverter
    {
        private static Dictionary<string, string> formatToSqlType = new()
        {
            ["PERCENT"] = "DECIMAL NOT NULL",
            ["NUMBER"] = "DECIMAL NOT NULL",
            ["CURRENCY"] = "MONEY NOT NULL",
            ["DATE_TIME"] = "DATETIME NOT NULL",
            ["DATE"] = "DATE NOT NULL",
            ["TEXT"] = "VARCHAR(MAX) NOT NULL",
        };

        /// <param name="format">Format of the Cell</param>
        /// <returns>T-Sql type of the format</returns>
        public static string GetSqlType(string format)
        {
            return formatToSqlType[format];
        }

        /// <summary>Changes cell value format to T-Sql format</summary>
        /// <param name="colFormats"></param>
        public static void FixFormatForSql(IList<IList<object>> values,
            string[] colFormats)
        {
            for (int i = 0; i < colFormats.Length; i++)
            {
                switch (colFormats[i])
                {
                    case "PERCENT":
                        RemoveCommasFromColumn(values, i);
                        RemovePercentSignFromColumn(values, i);
                        break;
                    case "NUMBER" or "CURRENCY":
                        RemoveCommasFromColumn(values, i);
                        break;
                    default:
                        break;
                }
            }
        }

        private static void RemoveCommasFromColumn(IList<IList<object>> values,
            int col)
        {
            for (int i = 1; i < values.Count; i++)
            {
                values[i][col] = values[i][col].ToString().Replace(",", "");
            }
        }
        private static void RemovePercentSignFromColumn(IList<IList<object>> values,
            int col)
        {
            for (int i = 1; i < values.Count; i++)
            {
                var str = values[i][col].ToString();
                values[i][col] = str.Substring(0, str.Length - 1);
            }
        }
    }
}
