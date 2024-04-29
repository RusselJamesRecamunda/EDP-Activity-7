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
    public partial class Form1 : Form
    {
        public static string LoggedInUserEmail { get; private set; }
        private string connectionString = "server=127.0.0.1; uid=root; pwd=Sparrowcakes212024; database=lawoffice";

        public Form1()
        {
            InitializeComponent();
            // Set the LinkBehavior property to remove the underline
            linkLabel1.LinkBehavior = LinkBehavior.NeverUnderline;

            // Set MaximizeBox and MinimizeBox to false
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string email = log_email.Text;
            string password = log_pass.Text;

            MySqlConnection conn = new MySqlConnection(connectionString);
            try
            {
                conn.Open();
                string sql = "SELECT userID, password, account_status FROM users WHERE email = @email";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@email", email);
                MySqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    int userId = Convert.ToInt32(reader["userID"]);
                    string storedPassword = reader["password"].ToString();
                    string accountStatus = reader["account_status"].ToString();
                    reader.Close(); // Close the DataReader here

                    if (password == storedPassword)
                    {
                        if (accountStatus == "Active")
                        {
                            // Update activity_status to "Active"
                            string updateSql = "UPDATE users SET activity_status = 'Online' WHERE userID = @userId";
                            MySqlCommand updateCmd = new MySqlCommand(updateSql, conn);
                            updateCmd.Parameters.AddWithValue("@userId", userId);
                            updateCmd.ExecuteNonQuery();

                            if (userId == 1 && email == "admin@gmail.com")
                            {
                                LoggedInUserEmail = email;
                                MessageBox.Show("Admin login successful!");
                                Dashboard adminDashboard = new Dashboard();
                                adminDashboard.Show();
                                this.Hide();
                            }
                            else
                            {
                                LoggedInUserEmail = email;
                                MessageBox.Show("User login successful!");

                                // Open the user dashboard and pass the email
                                user_dashboard userDashboard = new user_dashboard(email);
                                userDashboard.Show();
                                this.Hide();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Your account is inactive. Please contact support.");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid password. Please try again.");
                    }
                }
                else
                {
                    MessageBox.Show($"Invalid email ({email}). Please try again.");
                }
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



        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            verify Verify = new verify();
            Verify.Show();
            this.Hide();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
       
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}
