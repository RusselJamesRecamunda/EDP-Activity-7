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
using System.Xml.Linq;

namespace GUI
{
    public partial class profile : Form
    {
        private Handler dbHandler;
        private string userEmail;

        public profile(string email)
        {
            InitializeComponent();
            userEmail = email;

            // Connection string
            string connectionString = "server=127.0.0.1; uid=root; pwd=Sparrowcakes212024; database=lawoffice";

            // Create instance of Handler with connection string
            dbHandler = new Handler(connectionString);
            DisplayUserData();
        }

        private void DisplayUserData()
        {
            string query = $"SELECT * FROM users WHERE email = '{userEmail}'";
            DataTable userData = dbHandler.Read(query);

            if (userData.Rows.Count == 0)
            {
                MessageBox.Show("User data not found.");
                return;
            }

            DataRow row = userData.Rows[0];
            u_fname.Text = row["first_name"].ToString();
            u_mname.Text = row["middle_name"].ToString();
            u_lname.Text = row["last_name"].ToString();
            u_bdate.Text = row["birthdate"].ToString();
            u_email.Text = row["email"].ToString();
            u_phone.Text = row["phone"].ToString();
            u_address.Text = row["address"].ToString();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void f_name_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Get the updated values from text boxes
            string firstName = u_fname.Text;
            string middleName = u_mname.Text;
            string lastName = u_lname.Text;
            string birthDate = u_bdate.Text;
            string email = u_email.Text;
            string phone = u_phone.Text;
            string address = u_address.Text;

            // Update query
            string query = $"UPDATE users SET first_name = '{firstName}', middle_name = '{middleName}', last_name = '{lastName}', birthdate = '{birthDate}', phone = '{phone}', address = '{address}' WHERE email = '{userEmail}'";

            // Execute the update query
            int rowsAffected = dbHandler.Execute(query);

            if (rowsAffected > 0)
            {
                MessageBox.Show("User data updated successfully.");
            }
            else
            {
                MessageBox.Show("Failed to update user data.");
            }
        }

        private void u_fname_TextChanged(object sender, EventArgs e)
        {

        }
    }
    }