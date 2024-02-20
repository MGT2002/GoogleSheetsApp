using Microsoft.Data.SqlClient;

namespace GoogleSheetsApp
{
    internal class DataBaseManager(string connectionString)
    {
        public string ConnectionString { get; } = connectionString;

        /// <param name="query">T-Sql query</param>
        /// <returns>false if exception is thrown</returns>
        public bool ExecuteQuery(string query)
        {
            using SqlConnection myConn = new(ConnectionString);
            SqlCommand cmd = new(query, myConn);

            try
            {
                myConn.Open();
                cmd.ExecuteNonQuery();
                //Console.WriteLine("Success! Query executed");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return false;
            }

            return true;
        }
    }
}
