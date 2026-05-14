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

namespace afshin
{
    public partial class ClienInfoForm : Form
    {
        public ClienInfoForm()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox2.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string customerCodeText = textBox1.Text.Trim();
            string customerName = textBox2.Text.Trim();

            if (!int.TryParse(customerCodeText, out int customerCode))
            {
                MessageBox.Show("کد مشتری باید عدد باشد.");
                textBox1.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(customerName))
            {
                MessageBox.Show("لطفاً نام مشتری را وارد کنید.");
                textBox2.Focus();
                return;
            }

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
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    string query = @"
                INSERT INTO [Table2] ([کد مشتری], [نام مشتری])
                VALUES (?, ?)";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("?", customerCode);
                        cmd.Parameters.AddWithValue("?", customerName);

                        cmd.ExecuteNonQuery();
                    }
                }

                RefreshMainDataGridView();

                MessageBox.Show("مشتری با موفقیت اضافه شد.");

                textBox1.Clear();
                textBox2.Clear();
                textBox1.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در افزودن مشتری: " + ex.Message);
            }
        }

        private void RefreshMainDataGridView()
        {
            string dbPath = Properties.Settings.Default.LastDBPath;

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

            MainForm mainForm = Application.OpenForms["MainForm"] as MainForm;

            if (mainForm != null)
            {
                mainForm.dataGridView1.DataSource = null;
                mainForm.dataGridView1.AutoGenerateColumns = true;
                mainForm.dataGridView1.DataSource = dt;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            string customerCodeText = textBox1.Text.Trim();
            string customerName = textBox2.Text.Trim();

            if (!int.TryParse(customerCodeText, out int customerCode))
            {
                MessageBox.Show("کد مشتری باید عدد باشد.");
                textBox1.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(customerName))
            {
                MessageBox.Show("لطفاً نام مشتری را وارد کنید.");
                textBox2.Focus();
                return;
            }

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
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    string query = @"
                UPDATE [Table2]
                SET [نام مشتری] = ?
                WHERE [کد مشتری] = ?";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("?", customerName);
                        cmd.Parameters.AddWithValue("?", customerCode);

                        int rowsAffected = cmd.ExecuteNonQuery();

                        if (rowsAffected == 0)
                        {
                            MessageBox.Show("مشتری با این کد پیدا نشد.");
                            return;
                        }
                    }
                }

                MainForm mainForm = Application.OpenForms["MainForm"] as MainForm;

                if (mainForm != null)
                {
                    mainForm.RefreshTable2();
                }

                MessageBox.Show("اطلاعات مشتری با موفقیت ویرایش شد.");

                textBox1.Clear();
                textBox2.Clear();
                textBox1.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در ویرایش مشتری: " + ex.Message);
            }
        }
    }
}
