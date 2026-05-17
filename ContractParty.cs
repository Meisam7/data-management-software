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
    public partial class ContractParty : Form
    {
        public ContractParty()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string contractPartyName = textBox9.Text.Trim();

            if (string.IsNullOrWhiteSpace(contractPartyName))
            {
                MessageBox.Show("لطفاً نام دستگاه را وارد کنید.");
                textBox9.Focus();
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
                INSERT INTO [Table3] ([نام دستگاه])
                VALUES (?)";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("?", contractPartyName);
                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("نام دستگاه با موفقیت اضافه شد.");

                textBox9.Clear();
                textBox9.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در افزودن دستگاه: " + ex.Message);
            }

        }
    }
}
