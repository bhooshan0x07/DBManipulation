using System;
using System.Data.SqlClient;
using ClosedXML.Excel;

class DBManipulation
{
    static void Main(string[] args)
    {
        string excelPath = @"C:\Users\DELL\Downloads\data.xlsx";
        string connectionString = "Server=YourServerName;Database=YourDatabaseName;Trusted_Connection=True;";
        string tableName = "Location";
       
        Console.WriteLine("Initializing........");
        Console.WriteLine("Reading xlsx file: " + excelPath);

        // Read the Excel file
        // Temporarily enable IDENTITY_INSERT
        using (var workbook = new XLWorkbook(excelPath))
        {
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RowsUsed();
            int counter = 0;
            // comment/uncomment below 3 lines if you need to turn on indentity insert
            //Console.WriteLine("Enabling Identity Insert...");
            //SetIdentityInsert(connectionString, tableName, true);
            //Console.WriteLine("Enabled Identity Insert");
            try
            {
                // Iterate through the rows
                foreach (var row in rows.Skip(1)) // Assuming the first row is the header
                {
                    int id = int.Parse(row.Cell(1).GetValue<string>());
                    string column1 = row.Cell(2).GetValue<string>();
                    string column2 = row.Cell(3).GetValue<string>();
                    counter++;
                    // Insert data into SQL Server
                    InsertData(connectionString, column1, column2, tableName);
                }
                Console.WriteLine($"Total {counter} record(s) inserted.");
            }
            finally
            {

                // comment/uncomment below 3 lines if you need to turn on indentity insert

                // Disable IDENTITY_INSERT after the operation
                //Console.WriteLine("Disabling Identity Insert...");

                //SetIdentityInsert(connectionString, tableName, false);

                //Console.WriteLine("Disabled Identity Insert");
            }

        }
    }


    static void InsertData(string connectionString, string column1, string column2, string tblName)
    {
        string query = $"INSERT INTO {tblName} (LocationId, Name, Status) VALUES (@LocationId, @Name, @Status)";
        Console.WriteLine("Inserting data into DB");

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            SqlCommand command = new SqlCommand(query, connection);
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
