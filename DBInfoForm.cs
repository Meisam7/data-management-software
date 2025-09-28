using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;

namespace afshin
{

    public partial class DBInfoForm : Form
    {
        public DBInfoForm()
        {
            InitializeComponent();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select File";
            openFileDialog.Filter = "All Files (*.*)|*.*|Access Database File (*accdb*)|*.accdb";
            openFileDialog.ShowDialog();
            label2.Text = openFileDialog.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string dbPath = label2.Text;  // Assuming you use a TextBox to enter the DB path

            try
            {
                // Save the path in settings
                Properties.Settings.Default.LastDBPath = dbPath;
                Properties.Settings.Default.Save();

                // Set the path in DbHelper class (if needed)
                Class1.SetDatabasePath(dbPath);

                MessageBox.Show("Database connected successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.Close();  // Close the form after successful connection
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DBInfoForm_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(Properties.Settings.Default.LastDBPath))
            {
                label2.Text = Properties.Settings.Default.LastDBPath;  // Display the saved path
            }
            else
            {
                label2.Text = string.Empty;  // Empty if no saved path
            }
        }
    }
}
