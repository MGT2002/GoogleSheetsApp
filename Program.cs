using Google.Apis.Sheets.v4.Data;
using GoogleSheetsApp;
using System.Text;

#region InputUrl
Console.WriteLine("Input your Google sheets url");
string url;
while (!SheetManager.IsGoogleSheetsUrl(url = Console.ReadLine() ?? ""))
    Console.WriteLine("Wrong Input!!!");
Console.WriteLine("Getting data, please wait...");
#endregion

SheetManager sheetManager = new(url);
//ZZZ(18278) in sheets, but Ms Sql can handle only 4096 columns
ValueRange? response = sheetManager.GetData("A1:ZZ");

if (response?.Values == null)
{
    Console.WriteLine("No data found in the specified range.");
    return;
}
Console.WriteLine("Data has gotten!");

#region SQL Server solution
string connectionString = "Data Source=.\\SQLEXPRESS;Initial "
      + "Catalog=GoogleSheetsAppDB;Integrated Security=True;Encrypt=False";
DataBaseManager dataBaseManager = new(connectionString);

string[] colFormats = sheetManager.GetColFormats();
if (colFormats == null)
{
    Console.WriteLine("Sheet column format is unreachable!");
    return;
}
FormatToSqlConverter.FixFormatForSql(response.Values, colFormats);

#region Query Create Table
Console.WriteLine("Create table in db...");
StringBuilder query = new(
        @"IF OBJECT_ID (N'SheetTable', 'U') IS NOT NULL begin
	    drop table SheetTable
	    end
	    CREATE TABLE SheetTable (
		    ID INT PRIMARY KEY IDENTITY,"
    );
int length = response.Values[0].Count;
for (int i = 0; i < length; i++)
    query.Append($"_{response.Values[0][i].ToString()?.Replace(" ", "")}"
        + $" {FormatToSqlConverter.GetSqlType(colFormats[i])}, ");
query.Append(");");

if (!dataBaseManager.ExecuteQuery(query.ToString()))
    return;
Console.WriteLine("Created!");
#endregion

#region Query Insert Table
Console.WriteLine("Inserting table values...");
query.Clear();
query.Append("INSERT INTO SheetTable VALUES ");
for (int i = 1; i < response.Values.Count; i++)
{
    query.Append("(");
    for (int j = 0; j < length; j++)
        query.Append($"'{response.Values[i][j]}', ");
    query[^2] = ' ';
    query.Append("), ");
    //("('Value1A', 'Value2A', 'Value3A'),");
}
query[^2] = ';';

if (!dataBaseManager.ExecuteQuery(query.ToString()))
    return;
Console.WriteLine("Success!");
#endregion 

#endregion

Console.WriteLine("Press 'Enter' to exit.");
Console.Read();