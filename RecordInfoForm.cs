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
            string requestId = textBox4.Text.Trim();              // شناسه درخواست
            string contractParty = comboBox5.Text.Trim();         // نام دستگاه
            string province = comboBox2.Text.Trim();              // استان
            string havaleId = textBox5.Text.Trim();               // شناسه حواله
            string assistantName = textBox6.Text.Trim();          // معاونت
            string havaleNumberText = textBox11.Text.Trim();      // شماره حواله
            string havaleDate = textBox10.Text.Trim();            // تاریخ حواله
            string customerName = comboBox6.Text.Trim();          // نام مشتری
            string bitumenType = comboBox1.Text.Trim();           // نوع قیر
            string receiver = comboBox4.Text.Trim();              // گیرنده
            string havalePriceText = textBox15.Text.Trim();       // مبلغ حواله
            string allocationYearText = textBox1.Text.Trim();     // سال اعتبارات
            string havaleAmountText = textBox7.Text.Trim();       // مقدار حواله
            string taxPriceText = textBox12.Text.Trim();          // فی (با ارزش افزوده)
            string sendAmountText = textBox3.Text.Trim();         // ارسال
            string buyAmountText = textBox21.Text.Trim();         // خرید
            string finalAmountText = textBox8.Text.Trim();        // مقدار نهایی
            string carrierName = textBox22.Text.Trim();           // پیمانکار حمل
            string invoiceAmountText = textBox13.Text.Trim();     // مبلغ فاکتور
            string remainingHavale = textBox19.Text.Trim();       // مانده حواله
            string remainingInvoice = textBox16.Text.Trim();      // مانده فاکتور
            string description = textBox2.Text.Trim();            // توضیحات

            // ===============================
            // Required fields
            // ===============================

            if (string.IsNullOrWhiteSpace(contractParty))
            {
                MessageBox.Show("لطفاً نام دستگاه را وارد کنید.");
                comboBox5.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(province))
            {
                MessageBox.Show("لطفاً استان را وارد کنید.");
                comboBox2.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(havaleNumberText))
            {
                MessageBox.Show("لطفاً شماره حواله را وارد کنید.");
                textBox11.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(havaleDate))
            {
                MessageBox.Show("لطفاً تاریخ حواله را وارد کنید.");
                textBox10.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(customerName))
            {
                MessageBox.Show("لطفاً نام مشتری را وارد کنید.");
                comboBox6.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(bitumenType))
            {
                MessageBox.Show("لطفاً نوع قیر را وارد کنید.");
                comboBox1.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(receiver))
            {
                MessageBox.Show("لطفاً گیرنده را وارد کنید.");
                comboBox4.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(havaleAmountText))
            {
                MessageBox.Show("لطفاً مقدار حواله را وارد کنید.");
                textBox7.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(allocationYearText))
            {
                MessageBox.Show("لطفاً سال اعتبارات را وارد کنید.");
                textBox1.Focus();
                return;
            }

            // ===============================
            // Convert numeric fields
            // ===============================

            int? requestIdValue = ToNullableInt(requestId);
            int? havaleIdValue = ToNullableInt(havaleId);
            int? havaleNumberValue = ToNullableInt(havaleNumberText);
            int? allocationYearValue = ToNullableInt(allocationYearText);

            decimal? havalePriceValue = ToNullableDecimal(havalePriceText);
            decimal? havaleAmountValue = ToNullableDecimal(havaleAmountText);
            decimal? taxPriceValue = ToNullableDecimal(taxPriceText);
            decimal? sendAmountValue = ToNullableDecimal(sendAmountText);
            decimal? buyAmountValue = ToNullableDecimal(buyAmountText);
            decimal? finalAmountValue = ToNullableDecimal(finalAmountText);
            decimal? invoiceAmountValue = ToNullableDecimal(invoiceAmountText);

            // مقدار تهاتر
            decimal? tahatorAmountValue = ToNullableDecimal(textBox20.Text.Trim());

            // محاسبه مانده حواله
            decimal remainingHavaleValue =
                (finalAmountValue ?? 0)
                - ((sendAmountValue ?? 0) + (buyAmountValue ?? 0) + (tahatorAmountValue ?? 0));

            textBox19.Text = remainingHavaleValue.ToString();

            remainingHavale = textBox19.Text.Trim();

            if (havaleNumberValue == null)
            {
                MessageBox.Show("شماره حواله باید عدد باشد.");
                textBox11.Focus();
                return;
            }

            if (havaleAmountValue == null)
            {
                MessageBox.Show("مقدار حواله باید عدد باشد.");
                textBox7.Focus();
                return;
            }

            if (allocationYearValue == null)
            {
                MessageBox.Show("سال اعتبارات باید عدد باشد.");
                textBox1.Focus();
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
                INSERT INTO [Table1]
                (
                    [دستگاه طرف قرارداد],
                    [استان],
                    [سال اعتبارات],
                    [شناسه حواله],
                    [شناسه درخواست],
                    [معاونت],
                    [شماره حواله],
                    [تاریخ حواله],
                    [نام مشتری],
                    [نوع قیر],
                    [گیرنده],
                    [مبلغ حواله],
                    [مقدار حواله],
                    [فی (با ارزش افزوده)],
                    [ارسال],
                    [خرید],
                    [مقدار نهایی],
                    [پیمانکار حمل],
                    [مبلغ فاکتور],
                    [مانده حواله],
                    [مانده فاکتور],
                    [توضیحات]
                )
                VALUES
                (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                )";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        AddParam(cmd, contractParty);          // دستگاه طرف قرارداد
                        AddParam(cmd, province);               // استان
                        AddParam(cmd, allocationYearValue);    // سال اعتبارات
                        AddParam(cmd, havaleIdValue);          // شناسه حواله
                        AddParam(cmd, requestIdValue);         // شناسه درخواست
                        AddParam(cmd, assistantName);          // معاونت
                        AddParam(cmd, havaleNumberValue);      // شماره حواله
                        AddParam(cmd, havaleDate);             // تاریخ حواله
                        AddParam(cmd, customerName);           // نام مشتری
                        AddParam(cmd, bitumenType);            // نوع قیر
                        AddParam(cmd, receiver);               // گیرنده
                        AddParam(cmd, havalePriceValue);       // مبلغ حواله
                        AddParam(cmd, havaleAmountValue);      // مقدار حواله
                        AddParam(cmd, taxPriceValue);          // فی (با ارزش افزوده)
                        AddParam(cmd, sendAmountValue);        // ارسال
                        AddParam(cmd, buyAmountValue);         // خرید
                        AddParam(cmd, finalAmountValue);       // مقدار نهایی
                        AddParam(cmd, carrierName);            // پیمانکار حمل
                        AddParam(cmd, invoiceAmountValue);     // مبلغ فاکتور
                        AddParam(cmd, remainingHavaleValue);        // مانده حواله
                        AddParam(cmd, remainingInvoice);       // مانده فاکتور
                        AddParam(cmd, description);            // توضیحات

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("رکورد با موفقیت اضافه شد.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در افزودن رکورد: " + ex.Message);
            }
        }
        private int? ToNullableInt(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            text = text.Replace(",", "").Trim();

            if (int.TryParse(text, out int value))
                return value;

            return null;
        }

        private decimal? ToNullableDecimal(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            text = text.Replace(",", "").Trim();

            if (decimal.TryParse(text, out decimal value))
                return value;

            return null;
        }

        private void AddParam(OleDbCommand cmd, object value)
        {
            if (value == null || value == DBNull.Value)
                cmd.Parameters.AddWithValue("?", DBNull.Value);
            else
                cmd.Parameters.AddWithValue("?", value);
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
