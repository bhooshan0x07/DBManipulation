using System;
using System.Data.SqlClient;
using ClosedXML.Excel;

class DBManipulation
{
    static void Main(string[] args)
    {
        string excelPath = @"C:\Users\DELL\Downloads\data.xlsx";
        string connectionString = "Server=YourServerName;Database=YourDatabaseName;Trusted_Connection=True;";

        Console.WriteLine("Initializing........");
        Console.WriteLine("Reading xlsx file: " + excelPath);

        // Read the Excel file
        // Temporarily enable IDENTITY_INSERT
        using (var workbook = new XLWorkbook(excelPath))
        {
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RowsUsed();

            Console.WriteLine("Enabling Identity Insert...");
            SetIdentityInsert(connectionString, "Location", true);
            Console.WriteLine("Enabled Identity Insert");
            try
            {
                // Iterate through the rows
                foreach (var row in rows.Skip(1)) // Assuming the first row is the header
                {
                    int id = int.Parse(row.Cell(1).GetValue<string>());
                    string column1 = row.Cell(2).GetValue<string>();
                    string column2 = row.Cell(3).GetValue<string>();
                    string tableName = row.Cell(4).GetValue<string>();

                    // Insert data into SQL Server
                    InsertData(connectionString, id, column1, column2, tableName);
                }
            }
            finally
            {


                // Disable IDENTITY_INSERT after the operation
                Console.WriteLine("Disabling Identity Insert...");

                SetIdentityInsert(connectionString, "Location", false);

                Console.WriteLine("Disabled Identity Insert");
            }

        }
    }

    static void InsertData(string connectionString, int id, string column1, string column2, string tblName)
    {
        string tableName = tblName;
        string query = $"INSERT INTO {tableName} (Id, LocationId, Name, Status) VALUES (@Id, @LocationId, @Name, @Status)";
        Console.WriteLine("Inserting data into DB");

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            SqlCommand command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Id", id);
            command.Parameters.AddWithValue("@LocationId", column1);
            command.Parameters.AddWithValue("@Name", column2);
            command.Parameters.AddWithValue("@Status", 1);

            connection.Open();
            command.ExecuteNonQuery();
            connection.Close();
        }
    }

    static void SetIdentityInsert(string connectionString, string tableName, bool enable)
    {
        string query = enable ? $"SET IDENTITY_INSERT {tableName} ON;" : $"SET IDENTITY_INSERT {tableName} OFF;";

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            SqlCommand command = new SqlCommand(query, connection);

            connection.Open();
            command.ExecuteNonQuery();
            connection.Close();
        }
    }
}
