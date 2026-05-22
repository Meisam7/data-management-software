using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;


namespace afshin
{
    public partial class MainForm : Form
    {
        string LastDBPath = Properties.Settings.Default.LastDBPath;
        Form2 form2 = new Form2();
        DBInfoForm dBInfoForm = new DBInfoForm();
        ClienInfoForm clienInfoForm = new ClienInfoForm();
        ContractParty ContractParty=new ContractParty();
        private string currentTable = "";
        public MainForm()
        {
            InitializeComponent();
            dataGridView1.SelectionChanged += DataGridView1_SelectionChanged;
            dataGridView1.KeyDown += dataGridView1_KeyDown;
            dataGridView1.CellDoubleClick += dataGridView1_CellDoubleClick;


            dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
            dataGridView1.MultiSelect = true;
        }

        private async void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            if (currentTable != "Table1")
                return;

            string clickedColumnHeader = dataGridView1.Columns[e.ColumnIndex].HeaderText;
            string clickedColumnName = dataGridView1.Columns[e.ColumnIndex].Name;

            if (clickedColumnHeader != "ردیف" && clickedColumnName != "ردیف")
                return;

            object rowIdObj = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

            if (rowIdObj == null || rowIdObj == DBNull.Value)
            {
                MessageBox.Show("شماره ردیف معتبر نیست.");
                return;
            }

            if (!int.TryParse(rowIdObj.ToString(), out int rowId))
            {
                MessageBox.Show("شماره ردیف معتبر نیست.");
                return;
            }

            Form2 editForm = new Form2(rowId);

            if (editForm.ShowDialog() == DialogResult.OK)
            {
                // بعد از بروزرسانی، دوباره Table1 را نمایش بده
                await LoadTableToGridAsync("Table1");
            }
        }

        private async Task LoadTableToGridAsync(string tableName)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            string dbPath = Properties.Settings.Default.LastDBPath;

            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            {
                MessageBox.Show("مسیر دیتابیس معتبر نیست.");
                return;
            }

            string provider = Path.GetExtension(dbPath).ToLower() == ".mdb"
                ? "Microsoft.Jet.OLEDB.4.0"
                : "Microsoft.ACE.OLEDB.16.0";

            string connectionString =
                $@"Provider={provider};Data Source={dbPath};Persist Security Info=False;";

            DataTable dt = await Task.Run(() =>
            {
                DataTable table = new DataTable();

                using (OleDbConnection conn = new OleDbConnection(connectionString))
                using (OleDbDataAdapter da = new OleDbDataAdapter($"SELECT * FROM [{tableName}]", conn))
                {
                    da.Fill(table);
                }

                return table;
            });

            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = dt;

            currentTable = tableName;
        }
        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            UpdateSelectedSum();
        }
        private void UpdateSelectedSum()
        {
            decimal sum = 0m;
            int count = 0;

            foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
            {
                if (cell?.Value == null) continue;

                // Parse numbers using current culture (handles comma/decimal for your locale)
                if (decimal.TryParse(
                    cell.Value.ToString(),
                    NumberStyles.Any,
                    CultureInfo.CurrentCulture,
                    out var value))
                {
                    sum += value;
                    count++;
                }
            }

            // Show Sum (and Count if you like)
            toolStripStatusLabelSum.Text = (count > 0)
                ? $"جمع: {sum:N0}   تعداد: {count}"
                : "جمع: ۰   تعداد: ۰";

        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            toolStripStatusLabelSum.Text = "Sum: 0   Count: 0";
        }
        public MainForm(DataTable tableData)
        {
            InitializeComponent();
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = tableData;
        }
        //////////////////////////
        /// ////////////////////////////

        private void Form1_Load(object sender, EventArgs e)
        {
            // شروع برنامه: هیچ جدولی نمایش داده نشود
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            // شروع برنامه: هر دو بخش جستجو غیرفعال باشند
            groupBox1.Enabled = false;
            groupBox2.Enabled = false;

            currentTable = "";

            toolStripStatusLabelSum.Text = "جمع: ۰   تعداد: ۰";

            dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
            dataGridView1.MultiSelect = true;

            string dbPath = (Properties.Settings.Default.LastDBPath ?? "").Trim('"', ' ');

            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            {
                DBInfoForm form3 = new DBInfoForm();
                form3.ShowDialog();
            }
        }

        private async void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Delete)
                return;

            if (currentTable != "Table1")
            {
                MessageBox.Show("حذف فقط برای جدول حواله‌ها فعال است.");
                e.Handled = true;
                return;
            }

            // اگر کاربر فقط سلول انتخاب کرده، حذف رکورد انجام نشود
            if (dataGridView1.SelectedRows.Count == 0)
                return;

            if (!dataGridView1.Columns.Contains("ردیف"))
            {
                MessageBox.Show("ستون ردیف پیدا نشد.");
                return;
            }

            List<int> rowIds = new List<int>();

            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                if (row.IsNewRow)
                    continue;

                object rowIdObj = row.Cells["ردیف"].Value;

                if (rowIdObj != null && rowIdObj != DBNull.Value)
                {
                    if (int.TryParse(rowIdObj.ToString(), out int rowId))
                    {
                        rowIds.Add(rowId);
                    }
                }
            }

            if (rowIds.Count == 0)
            {
                MessageBox.Show("هیچ ردیف معتبری برای حذف پیدا نشد.");
                return;
            }

            DialogResult result = MessageBox.Show(
                $"آیا مطمئن هستید که می‌خواهید {rowIds.Count} رکورد را حذف کنید؟",
                "تأیید حذف",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (result != DialogResult.Yes)
                return;

            e.Handled = true;

            string dbPath = Properties.Settings.Default.LastDBPath;

            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            {
                MessageBox.Show("مسیر دیتابیس معتبر نیست.");
                return;
            }

            string provider = Path.GetExtension(dbPath).ToLower() == ".mdb"
                ? "Microsoft.Jet.OLEDB.4.0"
                : "Microsoft.ACE.OLEDB.16.0";

            string connectionString =
                $@"Provider={provider};Data Source={dbPath};Persist Security Info=False;";

            try
            {
                int deletedCount = 0;

                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    foreach (int rowId in rowIds)
                    {
                        string query = "DELETE FROM [Table1] WHERE [ردیف] = ?";

                        using (OleDbCommand cmd = new OleDbCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("?", rowId);
                            deletedCount += cmd.ExecuteNonQuery();
                        }
                    }
                }

                MessageBox.Show($"{deletedCount} رکورد با موفقیت حذف شد.");

                await LoadTableToGridAsync("Table1");
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در حذف رکوردها: " + ex.Message);
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void افزودنToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form2.ShowDialog();
        }

        private void ویرایشToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void پایگاهدادهToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dBInfoForm.ShowDialog();
        }

        private void اطToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void تنظیماتپایهToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void اطلاعاتمشتریانToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void رکورداصلیToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form2.ShowDialog();
        }

        private async void جدولاصلیToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                string dbPath = Properties.Settings.Default.LastDBPath;

                if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath))
                {
                    MessageBox.Show("Database path is not valid or set.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Class1.SetDatabasePath(dbPath);

                // Table3 = دستگاه‌های طرف قرارداد
                groupBox1.Enabled = false;
                groupBox2.Enabled = false;

                await Task.Run(() => Class1.LoadTableIntoGrid(dataGridView1, "Table3"));

                currentTable = "Table3";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void جدولمشتریانToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                string dbPath = Properties.Settings.Default.LastDBPath;

                if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath))
                {
                    MessageBox.Show("Database path is not valid or set.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Class1.SetDatabasePath(dbPath);

                // Table2 = جدول مشتریان
                groupBox1.Enabled = false;
                groupBox2.Enabled = true;

                await Task.Run(() => Class1.LoadTableIntoGrid(dataGridView1, "Table2"));

                currentTable = "Table2";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void خروجیToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0) { MessageBox.Show("No data."); return; }

            var sfd = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "Export.csv" };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            var sb = new StringBuilder();

            // headers (only visible columns)
            for (int c = 0; c < dataGridView1.Columns.Count; c++)
                if (dataGridView1.Columns[c].Visible)
                    sb.Append(E(dataGridView1.Columns[c].HeaderText)).Append(',');

            if (sb.Length > 0) sb.Length--; // trim last comma
            sb.AppendLine();

            // rows
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;
                var line = new StringBuilder();
                for (int c = 0; c < dataGridView1.Columns.Count; c++)
                    if (dataGridView1.Columns[c].Visible)
                        line.Append(E(row.Cells[c].Value?.ToString())).Append(',');
                if (line.Length > 0) line.Length--;
                sb.AppendLine(line.ToString());
            }

            File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true)); // true = add BOM
            MessageBox.Show("Exported.");
        }
        private string E(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void افزودنToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            clienInfoForm.ShowDialog();
        }

        private async void فهرستمشتریانToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                // Clear DataGridView before loading new data
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                // Retrieve the path from settings and load the data
                string dbPath = Properties.Settings.Default.LastDBPath;

                if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath))
                {
                    MessageBox.Show("Database path is not valid or set.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Set the database path in Class1
                Class1.SetDatabasePath(dbPath);


                // Disable GroupBox1 and enable groupbox2
                groupBox1.Enabled = false;  // Disable GroupBox2
                groupBox2.Enabled = true;  // Enable GroupBox1

                // Load Table2 asynchronously into the grid
                await Task.Run(() => Class1.LoadTableIntoGrid(dataGridView1, "Table2"));

                // Update the flag to track that Table2 is loaded
                currentTable = "Table2";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource is DataTable dt)
            {
                string search = textBox1.Text.Replace("'", "''"); // escape quotes

                if (string.IsNullOrWhiteSpace(search))
                {
                    dt.DefaultView.RowFilter = ""; // show all
                }
                else
                {
                    // Build filter string for all columns
                    var filters = new List<string>();
                    foreach (DataColumn col in dt.Columns)
                    {
                        // Only for string-like columns
                        if (col.DataType == typeof(string))
                        {
                            filters.Add($"[{col.ColumnName}] LIKE '%{search}%'");
                        }
                        else
                        {
                            // For numeric/date columns, convert to string
                            filters.Add($"CONVERT([{col.ColumnName}], System.String) LIKE '%{search}%'");
                        }
                    }

                    dt.DefaultView.RowFilter = string.Join(" OR ", filters);
                }
            }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            // Get the DataTable bound to the DataGridView
            if (dataGridView1.DataSource is DataTable dt)
            {
                string filterText = textBox2.Text.Replace("'", "''"); // escape single quotes

                // Apply filter on the column دستگاه طرف قرارداد
                if (string.IsNullOrWhiteSpace(filterText))
                    dt.DefaultView.RowFilter = ""; // show all rows
                else
                    dt.DefaultView.RowFilter = $"[دستگاه طرف قرارداد] LIKE '%{filterText}%'";
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource is DataTable dt)
            {
                string search = textBox3.Text.Replace("'", "''");

                if (string.IsNullOrWhiteSpace(search))
                    dt.DefaultView.RowFilter = ""; // Show all rows
                else
                    dt.DefaultView.RowFilter = $"CONVERT([شماره حواله], System.String) LIKE '%{search}%'";
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource is DataTable dt)
            {
                string selectedValue = comboBox1.SelectedItem?.ToString().Replace("'", "''");

                if (string.IsNullOrWhiteSpace(selectedValue))
                    dt.DefaultView.RowFilter = ""; // Show all rows
                else
                    dt.DefaultView.RowFilter = $"[استان] = '{selectedValue}'";
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource is DataTable dt)
            {
                string search = textBox9.Text.Replace("'", "''");

                if (string.IsNullOrWhiteSpace(search))
                    dt.DefaultView.RowFilter = ""; // Show all rows
                else
                    dt.DefaultView.RowFilter = $"[نام مشتری] LIKE '%{search}%'";
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource is DataTable dt)
            {
                string search = textBox7.Text.Replace("'", "''");

                if (string.IsNullOrWhiteSpace(search))
                    dt.DefaultView.RowFilter = ""; // Show all rows
                else
                    dt.DefaultView.RowFilter = $"[گیرنده(محل تخلیه)] LIKE '%{search}%'";
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource is DataTable dt)
            {
                string search = textBox8.Text.Replace("'", "''");

                if (string.IsNullOrWhiteSpace(search))
                    dt.DefaultView.RowFilter = ""; // Show all rows
                else
                    dt.DefaultView.RowFilter = $"CONVERT([مقدار حواله], System.String) LIKE '%{search}%'";
            }
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource is DataTable dt)
            {
                string search = textBox10.Text.Replace("'", "''");

                if (string.IsNullOrWhiteSpace(search))
                    dt.DefaultView.RowFilter = ""; // Show all rows
                else
                    dt.DefaultView.RowFilter = $"[وضعیت فاکتور] LIKE '%{search}%'";
            }
        }

        public void RefreshTable2()
        {
            string dbPath = Properties.Settings.Default.LastDBPath;

            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            {
                MessageBox.Show("مسیر دیتابیس معتبر نیست.");
                return;
            }

            string provider = Path.GetExtension(dbPath).ToLower() == ".mdb"
                ? "Microsoft.Jet.OLEDB.4.0"
                : "Microsoft.ACE.OLEDB.16.0";

            string connectionString =
                $@"Provider={provider};Data Source={dbPath};Persist Security Info=False;";

            DataTable dt = new DataTable();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            using (OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Table2]", conn))
            {
                da.Fill(dt);
            }

            dataGridView1.DataSource = null;
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = dt;

            currentTable = "Table2";
        }

        private void افزودنToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            ContractParty.ShowDialog();
        }

        private async void فهرستToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                // Clear DataGridView before loading new data
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                // Retrieve the path from settings and load the data
                string dbPath = Properties.Settings.Default.LastDBPath;

                if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath))
                {
                    MessageBox.Show("Database path is not valid or set.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Set the database path in Class1
                Class1.SetDatabasePath(dbPath);

                // Disable GroupBox2 and enable groupbox1
                groupBox2.Enabled = false;  // Disable GroupBox2
                groupBox1.Enabled = true;  // Enable GroupBox1

                // Load Table1 asynchronously into the grid
                await Task.Run(() => Class1.LoadTableIntoGrid(dataGridView1, "Table3"));

                currentTable = "Table3";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void جدولحوالههاToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                string dbPath = Properties.Settings.Default.LastDBPath;

                if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath))
                {
                    MessageBox.Show("Database path is not valid or set.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Class1.SetDatabasePath(dbPath);

                // Table1 = جدول حواله‌ها
                groupBox1.Enabled = true;
                groupBox2.Enabled = false;

                await Task.Run(() => Class1.LoadTableIntoGrid(dataGridView1, "Table1"));

                currentTable = "Table1";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetSearchBoxesForTable(string tableName)
        {
            if (tableName == "Table1")
            {
                groupBox1.Enabled = true;
                groupBox2.Enabled = false;
            }
            else if (tableName == "Table2")
            {
                groupBox1.Enabled = false;
                groupBox2.Enabled = true;
            }
            else
            {
                groupBox1.Enabled = false;
                groupBox2.Enabled = false;
            }
        }
    }
}
