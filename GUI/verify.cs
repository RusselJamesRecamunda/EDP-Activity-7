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
    public partial class verify : Form
    {
        private string connectionString = "server=127.0.0.1; uid=root; pwd=Sparrowcakes212024; database=lawoffice";
        public verify()
        {
            InitializeComponent();
            this.MinimizeBox = false;
            this.MaximizeBox = false;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string email = v_email.Text.Trim();
            string newPassword = o_password.Text;
            string confirmPassword = rw_password.Text;

            // Check if the new password and confirm password match
            if (newPassword != confirmPassword)
            {
                MessageBox.Show("Passwords do not match. Please try again.");
                return;
            }

            // Verify the email and update password
            if (VerifyAndUpdatePassword(email, newPassword))
            {
                MessageBox.Show("Password recovered successfully!");
            }
            else
            {
                MessageBox.Show("Failed to update password. Email not found or invalid.");
            }
        }

        private bool VerifyAndUpdatePassword(string email, string newPassword)
        {
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    connection.Open();

                    // Check if the email exists in the users table
                    string query = "SELECT COUNT(*) FROM users WHERE email = @Email";
                    using (MySqlCommand command = new MySqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Email", email);
                        int count = Convert.ToInt32(command.ExecuteScalar());
                        if (count == 0)
                            return false; // Email not found
                    }

                    // Update the password for the given email
                    string updateQuery = "UPDATE users SET password = @NewPassword WHERE email = @Email";
                    using (MySqlCommand updateCommand = new MySqlCommand(updateQuery, connection))
                    {
                        updateCommand.Parameters.AddWithValue("@NewPassword", newPassword);
                        updateCommand.Parameters.AddWithValue("@Email", email);
                        updateCommand.ExecuteNonQuery();
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                return false;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form1 Login = new Form1();
            Login.Show();
            this.Hide();
        }
    }
}
