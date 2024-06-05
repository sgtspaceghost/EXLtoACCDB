using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using OfficeOpenXml;
using System.Linq;

namespace ExcelToAccess
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get user input
            string excelFile = GetExcelFileFromUser();
            string range = GetRangeFromUser();
            string directory = GetDirectoryFromUser();
            string databaseName = GetDatabaseNameFromUser();
            string sortingCriteria = GetSortingCriteriaFromUser();

            // Read Excel data
            DataTable data = ReadExcelData(excelFile, range);

            // Determine data types and read surrounding cells
            DataTable enrichedData = EnrichDataWithSurroundingCells(data, range, excelFile);

            // Sort Data
            DataTable sortedData = SortData(enrichedData, sortingCriteria);

            // Create Access database
            string connectionString = CreateAccessDatabase(directory, databaseName);

            // Export data to Access database
            ExportDataToAccessDatabase(sortedData, connectionString);

            Console.WriteLine("Data export to Access database completed successfully.");
        }

        static string GetExcelFileFromUser()
        {
            Console.Write("Enter the path to the Excel file: ");
            return Console.ReadLine();
        }

        static string GetRangeFromUser()
        {
            Console.Write("Enter the Excel range (e.g., A1:D10): ");
            return Console.ReadLine();
        }

        static string GetDirectoryFromUser()
        {
            Console.Write("Enter the directory to save the Access database: ");
            return Console.ReadLine();
        }

        static string GetDatabaseNameFromUser()
        {
            Console.Write("Enter the name of the Access database: ");
            return Console.ReadLine();
        }

        static string GetSortingCriteriaFromUser()
        {
            Console.Write("Enter the sorting criteria (e.g., column name): ");
            return Console.ReadLine();
        }

        static DataTable ReadExcelData(string excelFile, string range)
        {
            DataTable data = new DataTable();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFile)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                var cells = worksheet.Cells[range];
                bool headerProcessed = false;

                foreach (var cell in cells)
                {
                    if (!headerProcessed)
                    {
                        data.Columns.Add(cell.Text);
                    }
                }

                headerProcessed = true;

                foreach (var row in cells.GroupBy(c => c.Start.Row))
                {
                    DataRow dataRow = data.NewRow();
                    foreach (var cell in row)
                    {
                        dataRow[cell.Start.Column - 1] = cell.Text;
                    }
                    data.Rows.Add(dataRow);
                }
            }

            return data;
        }

        static DataTable EnrichDataWithSurroundingCells(DataTable data, string range, string excelFile)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFile)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                var cells = worksheet.Cells[range];

                foreach (DataRow row in data.Rows)
                {
                    foreach (DataColumn column in data.Columns)
                    {
                        int rowIndex = data.Rows.IndexOf(row) + 1;
                        int colIndex = data.Columns.IndexOf(column) + 1;

                        // Read surrounding cells
                        var cell = worksheet.Cells[rowIndex, colIndex];
                        var leftCell = worksheet.Cells[rowIndex, colIndex - 1].Value?.ToString();
                        var rightCell = worksheet.Cells[rowIndex, colIndex + 1].Value?.ToString();
                        var topCell = worksheet.Cells[rowIndex - 1, colIndex].Value?.ToString();
                        var bottomCell = worksheet.Cells[rowIndex + 1, colIndex].Value?.ToString();

                        // Enrich the data row with information from surrounding cells
                        row["LeftCell"] = leftCell ?? "";
                        row["RightCell"] = rightCell ?? "";
                        row["TopCell"] = topCell ?? "";
                        row["BottomCell"] = bottomCell ?? "";
                    }
                }
            }

            return data;
        }

        static DataTable SortData(DataTable data, string sortingCriteria)
        {
            DataView dv = data.DefaultView;
            dv.Sort = sortingCriteria;
            return dv.ToTable();
        }

        static string CreateAccessDatabase(string directory, string databaseName)
        {
            string dbPath = Path.Combine(directory, $"{databaseName}.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath}";

            ADOX.Catalog catalog = new ADOX.Catalog();
            catalog.Create(connectionString);

            return connectionString;
        }

        static void ExportDataToAccessDatabase(DataTable data, string connectionString)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Create table
                string createTableQuery = GenerateCreateTableQuery(data);
                OleDbCommand createTableCommand = new OleDbCommand(createTableQuery, connection);
                createTableCommand.ExecuteNonQuery();

                // Insert data
                foreach (DataRow row in data.Rows)
                {
                    string insertQuery = GenerateInsertQuery(data, row);
                    OleDbCommand insertCommand = new OleDbCommand(insertQuery, connection);
                    insertCommand.ExecuteNonQuery();
                }
            }
        }

        static string GenerateCreateTableQuery(DataTable data)
        {
            string query = "CREATE TABLE Data (";
            foreach (DataColumn column in data.Columns)
            {
                query += $"[{column.ColumnName}] TEXT,";
            }
            query = query.TrimEnd(',') + ")";
            return query;
        }

        static string GenerateInsertQuery(DataTable data, DataRow row)
        {
            string query = "INSERT INTO Data (";
            foreach (DataColumn column in data.Columns)
            {
                query += $"[{column.ColumnName}],";
            }
            query = query.TrimEnd(',') + ") VALUES (";
            foreach (DataColumn column in data.Columns)
            {
                query += $"'{row[column.ColumnName]}',";
            }
            query = query.TrimEnd(',') + ")";
            return query;
        }
    }
}
