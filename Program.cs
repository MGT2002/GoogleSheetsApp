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

#region Query Create Table
Console.WriteLine("Create table in db...");
StringBuilder query = new(
        @"IF OBJECT_ID (N'SheetTable', 'U') IS NOT NULL begin
	    drop table SheetTable
	    end
	    CREATE TABLE SheetTable (
		    ID INT PRIMARY KEY IDENTITY,"
    );
int maxLength = response.Values.Max(list => list.Count);
for (int i = 0; i < maxLength; i++)
    query.Append($"Col{i + 1} varchar(max) not null, ");
query.Append(");");

dataBaseManager.ExecuteQuery(query.ToString());
Console.WriteLine("Created!");
#endregion

#region Query Insert Table
Console.WriteLine("Inserting table values...");
query.Clear();
query.Append("INSERT INTO SheetTable VALUES ");
for (int i = 0; i < response.Values.Count; i++)
{
    query.Append("(");
    int j;
    for (j = 0; j < response.Values[i].Count; j++)
        query.Append($"'{response.Values[i][j]}', ");
    for (; j < maxLength; j++)
        query.Append("'', ");
    query[^2] = ' ';
    query.Append("), ");
    //("('Value1A', 'Value2A', 'Value3A', '', '', ...),");
}
query[^2] = ';';

dataBaseManager.ExecuteQuery(query.ToString());
Console.WriteLine("Success!");
#endregion 

#endregion

Console.WriteLine("Press 'Enter' to exit.");
Console.Read();