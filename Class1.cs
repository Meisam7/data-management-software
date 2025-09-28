using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace afshin
{
    public class Class1
    {
        // Persisted for reuse
        private static string _dbPath;

        public static void SetDatabasePath(string dbPath)
        {
            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
                throw new FileNotFoundException("Database file not found.", dbPath);

            _dbPath = dbPath;
        }

        private static string BuildConnString()
        {
            if (string.IsNullOrEmpty(_dbPath))
                throw new InvalidOperationException("Database path not set. Call SetDatabasePath first.");

            var ext = Path.GetExtension(_dbPath).ToLowerInvariant();
            string provider = (ext == ".mdb") ? "Microsoft.Jet.OLEDB.4.0" : "Microsoft.ACE.OLEDB.16.0";
            return $@"Provider={provider};Data Source={_dbPath};Persist Security Info=False;";
        }

        public static DataTable GetTable(string tableName = "Table1")
        {
            var dt = new DataTable();
            using (var conn = new OleDbConnection(BuildConnString()))
            {
                var da = new OleDbDataAdapter($@"SELECT * FROM [{tableName}]", conn);
                conn.Open();
                da.Fill(dt);  // Fill the DataTable
                conn.Close(); // Explicitly close the connection
            }
            return dt;
        }

        // Convenience: clear and load directly into a grid
        public static void LoadTableIntoGrid(DataGridView grid, string tableName = "Table1")
        {
            if (grid == null) return;

            // Run the database operation on a background thread (asynchronously)
            var dt = GetTable(tableName);

            // Update the grid on the UI thread once the data is retrieved
            grid.Invoke((Action)(() =>
            {
                // Ensure the DataGridView is cleared before loading new data
                grid.DataSource = null;
                grid.Rows.Clear();
                grid.Columns.Clear();

                // Set properties for the grid
                grid.AutoGenerateColumns = true;
                grid.DataSource = dt;
            }));
        }
    }
}
