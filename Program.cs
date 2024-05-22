using System;
using System.Data.SqlClient;
using ClosedXML.Excel;

class DBManipulation
{
    static void Main(string[] args)
    {
        string excelPath = @"C:\Path\To\data.xlsx";
        string connectionString = "Server=YourServerName;Database=YourDatabaseName;Trusted_Connection=True;";

        // Read the Excel file
        using (var workbook = new XLWorkbook(excelPath))
        {
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RowsUsed();

            // Iterate through the rows
            foreach (var row in rows.Skip(1)) // Assuming the first row is the header
            {
                int id = int.Parse(row.Cell(1).GetValue<string>());
                string column1 = row.Cell(2).GetValue<string>();
                string column2 = row.Cell(3).GetValue<string>();

                // Insert data into SQL Server
                InsertData(connectionString, id, column1, column2);
            }
        }
    }

    static void InsertData(string connectionString, int id, string column1, string column2)
    {
        string query = "INSERT INTO TargetTable (ID, Column1, Column2) VALUES (@ID, @Column1, @Column2)";

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            SqlCommand command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@ID", id);
            command.Parameters.AddWithValue("@Column1", column1);
            command.Parameters.AddWithValue("@Column2", column2);

            connection.Open();
            command.ExecuteNonQuery();
            connection.Close();
        }
    }
}
