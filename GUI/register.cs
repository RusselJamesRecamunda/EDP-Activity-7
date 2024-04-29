using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace GUI
{
    public partial class register : Form
    {
        private string LoggedInUserEmail;
        private string connectionString = "server=127.0.0.1; uid=root; pwd=Sparrowcakes212024; database=lawoffice";

        public register(string email)
        {
            InitializeComponent();
            LoggedInUserEmail = email;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
        }

        private void register_Load(object sender, EventArgs e)
        {
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString = "server=127.0.0.1;uid=root;pwd=Sparrowcakes212024;database=lawoffice";
            Handler dbHandler = new Handler(connectionString);

            // Example of CRUD operations
            // Insert

            string first_name = f_name.Text;
            string middle_name = m_name.Text;
            string last_name = l_name.Text;
            string birthdate = b_date.Text;
            string phone = txb_phone.Text;
            string address = txb_address.Text;
            string email = txb_email.Text;
            string password = n_password.Text;

            // Perform input validation if necessary

            // Insert user data into the database
            string query = $"INSERT INTO users (first_name, middle_name, last_name, birthdate, phone, address, email, password) " +
                $"VALUES ('{first_name}', '{middle_name}', '{last_name}', '{birthdate}', '{phone}', '{address}', '{email}', '{password}')";

            // Check if the query was successful
            try
            {
                // Execute the query
                dbHandler.Create(query);
                MessageBox.Show("Registered Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Registration failed. Error: {ex.Message}");
            }
        }

     

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void n_password_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void c_password_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void c_password_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void n_password_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void pictureBox15_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            Dashboard adm_dashboard = new Dashboard();
            adm_dashboard.Show();
            this.Hide();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            try
            {
                conn.Open();

                // Get the user ID of the currently logged-in user
                string email = LoggedInUserEmail;
                string getUserIdSql = "SELECT userID FROM users WHERE email = @email";
                MySqlCommand getUserIdCmd = new MySqlCommand(getUserIdSql, conn);
                getUserIdCmd.Parameters.AddWithValue("@email", email);
                int userId = Convert.ToInt32(getUserIdCmd.ExecuteScalar());

                // Update activity_status to "Inactive"
                string updateActivitySql = "UPDATE users SET activity_status = 'Offline' WHERE userID = @userId";
                MySqlCommand updateActivityCmd = new MySqlCommand(updateActivitySql, conn);
                updateActivityCmd.Parameters.AddWithValue("@userId", userId);
                updateActivityCmd.ExecuteNonQuery();

                // Clear the logged-in user email
                LoggedInUserEmail = null;

                // Show login form again
                Form1 loginForm = new Form1();
                loginForm.Show();
                this.Close(); // Close current register form
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void b_date_ValueChanged(object sender, EventArgs e)
        {

        }
    }
    }

