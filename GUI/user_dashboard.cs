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

namespace GUI
{
    public partial class user_dashboard: Form
    {
        private string LoggedInUserEmail;
        private string connectionString = "server=127.0.0.1; uid=root; pwd=Sparrowcakes212024; database=lawoffice";
        public user_dashboard(string email)
        {
            InitializeComponent();
            LoggedInUserEmail = email;
        }

        private void user_dashboard_Load(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            About aboutIS = new About();
            aboutIS.Show();
            this.Hide();
        }

        private void btn_reports_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            // Get the logged-in user's email
            string email = Form1.LoggedInUserEmail;

            // Open the profile form and pass the user's email
            profile userProfile = new profile(email);
            userProfile.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            try
            {
                conn.Open();

                // Get the user ID of the currently logged-in user
                string email = LoggedInUserEmail; // Use the loggedInUserEmail field
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
                this.Close(); // Close current dashboard form
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

    }
}
