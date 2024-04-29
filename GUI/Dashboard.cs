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
    public partial class Dashboard : Form
    {
        private Handler dbHandler;
        private DataGridViewComboBoxColumn comboBoxColumn;

        public Dashboard()
        {
            InitializeComponent();
            // Initialize the database handler with your connection string
            string connectionString = "server=127.0.0.1; uid=root; pwd=Sparrowcakes212024; database=lawoffice";
            dbHandler = new Handler(connectionString);

            // Create the ComboBox column for status only once during initialization
            comboBoxColumn = new DataGridViewComboBoxColumn();
            comboBoxColumn.HeaderText = "Status";
            comboBoxColumn.Name = "statusColumn";
            comboBoxColumn.Items.AddRange("Active", "Inactive");

            // Add the ComboBox column to the DataGridView if it doesn't exist
            if (!dataGridView1.Columns.Contains("statusColumn"))
            {
                dataGridView1.Columns.Add(comboBoxColumn);
            }
        }

        private void Dashboard_Load(object sender, EventArgs e)
        {
            // Call the method to load data into DataGridView
            LoadUserData();
        }

        private void LoadUserData()
        {
            try
            {
                // SQL query to select data from the "users" table
                string userDataQuery = "SELECT userID, first_name, middle_name, last_name, " +
                                       "birthdate, phone, address, email, " +
                                       "created_time, activity_status, account_status FROM users";

                // Call the Read method of dbHandler to retrieve user data
                DataTable userData = dbHandler.Read(userDataQuery);

                // Bind the DataGridView to the DataTable
                dataGridView1.DataSource = userData;

                // Iterate through each row in the DataGridView and set the value of the ComboBox column
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Get the "account_status" cell value
                    object accountStatusObj = row.Cells["account_status"].Value;

                    // Check if the value is not null before accessing it
                    if (accountStatusObj != null)
                    {
                        string accountStatus = accountStatusObj.ToString();

                        // Check if the value is one of the valid options in the ComboBox
                        if (comboBoxColumn.Items.Contains(accountStatus))
                        {
                            // Set the value of the ComboBox cell
                            DataGridViewComboBoxCell comboBoxCell = row.Cells["statusColumn"] as DataGridViewComboBoxCell;
                            comboBoxCell.Value = accountStatus;
                        }
                        else
                        {
                            // If the value is not valid, set it to the first item in the ComboBox
                            DataGridViewComboBoxCell comboBoxCell = row.Cells["statusColumn"] as DataGridViewComboBoxCell;
                            comboBoxCell.Value = comboBoxColumn.Items[0];
                        }
                    }
                }

                // Hide the "account_status" column
                dataGridView1.Columns["account_status"].Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void panel12_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label3_Click_1(object sender, EventArgs e)
        {

        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void vScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {

        }
        

        private void btn_grid_Click(object sender, EventArgs e)
        {
            try
            {
                // Iterate through each row in the DataGridView
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Get the user ID from the current row
                    int? userID = row.Cells["userID"].Value as int?;

                    // Get the status from the ComboBox cell
                    string newStatus = row.Cells["statusColumn"].Value?.ToString();

                    // Ensure both userID and newStatus are not null before proceeding
                    if (userID != null && newStatus != null)
                    {
                        // Update the user status in the database
                        string updateQuery = $"UPDATE users SET account_status = '{newStatus}' WHERE userID = {userID}";
                        dbHandler.Execute(updateQuery);
                    }
                }

                // Refresh the DataGridView to reflect the updated data
                LoadUserData();
                MessageBox.Show("User status updated successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating user status: " + ex.Message);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                // Get the userID entered in the textBox2
                int userID;
                if (!int.TryParse(textBox2.Text, out userID))
                {
                    MessageBox.Show("Please enter a valid numeric userID.");
                    return;
                }

                // SQL query to select data from the "users" table for the specific userID
                string userDataQuery = $"SELECT * FROM users WHERE userID = {userID}";

                // Call the Read method of dbHandler to retrieve user data
                DataTable userData = dbHandler.Read(userDataQuery);

                // Bind the DataGridView to the DataTable
                dataGridView1.DataSource = userData;

                // Hide the "account_status" column if it's not already hidden
                dataGridView1.Columns["account_status"].Visible = false;

                if (userData.Rows.Count == 0)
                {
                    MessageBox.Show("No user found with the entered userID.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void btn_dashboard_Click(object sender, EventArgs e)
        {
            Dashboard dashboard = new Dashboard();
            dashboard.Show();
            this.Hide();
        }

        private void btn_register_Click(object sender, EventArgs e)
        {
            string email = Form1.LoggedInUserEmail;
            register Register = new register(email);
            Register.Show();
            this.Hide();
        }

        private void btn_book_Click(object sender, EventArgs e)
        {

        }

        private void btn_notaryFee_Click(object sender, EventArgs e)
        {

        }

        private void btn_reports_Click(object sender, EventArgs e)
        {
            admin_reporting AdminReports = new admin_reporting();
            AdminReports.Show();
            this.Hide();
        }

        private void btn_about_Click(object sender, EventArgs e)
        {
            About aboutIS = new About();
            aboutIS.Show();
            this.Hide();
        }

        private void btn_logout_Click(object sender, EventArgs e)
        {
            Form1 Login = new Form1();
            Login.Show();
            this.Hide();
        }
    }
}
