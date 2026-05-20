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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void LoadCustomerNamesIntoComboBox()
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

            try
            {
                comboBox4.BeginUpdate();
                comboBox6.BeginUpdate();

                comboBox4.Items.Clear();
                comboBox6.Items.Clear();

                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    string query = "SELECT [نام مشتری] FROM [Table2] ORDER BY [نام مشتری]";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader["نام مشتری"] != DBNull.Value)
                            {
                                string customerName = reader["نام مشتری"].ToString().Trim();

                                if (!string.IsNullOrWhiteSpace(customerName))
                                {
                                    comboBox4.Items.Add(customerName);
                                    comboBox6.Items.Add(customerName);
                                }
                            }
                        }
                    }
                }

                comboBox4.SelectedIndex = -1;
                comboBox6.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در خواندن نام مشتری‌ها: " + ex.Message);
            }
            finally
            {
                comboBox4.EndUpdate();
                comboBox6.EndUpdate();
            }
        }

        private void LoadGovernmentOrganizationsIntoComboBox5()
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

            try
            {
                comboBox5.BeginUpdate();
                comboBox5.Items.Clear();

                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    string query = "SELECT [نام دستگاه] FROM [Table3] ORDER BY [نام دستگاه]";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader["نام دستگاه"] != DBNull.Value)
                            {
                                string deviceName = reader["نام دستگاه"].ToString().Trim();

                                if (!string.IsNullOrWhiteSpace(deviceName))
                                {
                                    comboBox5.Items.Add(deviceName);
                                }
                            }
                        }
                    }
                }

                comboBox5.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در خواندن نام دستگاه‌ها: " + ex.Message);
            }
            finally
            {
                comboBox5.EndUpdate();
            }
        }


        private void LoadBitumenNamesIntoComboBox1()
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

            try
            {
                comboBox1.BeginUpdate();
                comboBox1.Items.Clear();

                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    string query = "SELECT [نام قیر] FROM [Table4]";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader["نام قیر"] != DBNull.Value)
                            {
                                string bitumenName = reader["نام قیر"].ToString().Trim();

                                if (!string.IsNullOrWhiteSpace(bitumenName))
                                {
                                    comboBox1.Items.Add(bitumenName);
                                }
                            }
                        }
                    }
                }

                comboBox1.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در خواندن نام قیرها: " + ex.Message);
            }
            finally
            {
                comboBox1.EndUpdate();
            }
        }

        private void SetupComboBox(ComboBox comboBox)
        {
            comboBox.ForeColor = Color.Black;
            comboBox.BackColor = Color.White;
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox.DrawMode = DrawMode.Normal;
            comboBox.IntegralHeight = false;
            comboBox.DropDownHeight = 120;
            comboBox.RightToLeft = RightToLeft.Yes;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            comboBox2.ForeColor = Color.Black;
            comboBox2.BackColor = Color.White;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DrawMode = DrawMode.Normal;
            comboBox2.IntegralHeight = false;
            comboBox2.DropDownHeight = 120;
            //------------------------------------------------------
            SetupComboBox(comboBox5);
            SetupComboBox(comboBox4);
            SetupComboBox(comboBox6);
            SetupComboBox(comboBox1);

            LoadGovernmentOrganizationsIntoComboBox5();
            LoadCustomerNamesIntoComboBox();
            LoadBitumenNamesIntoComboBox1();

        }
    }
}
