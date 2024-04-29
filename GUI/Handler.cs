using System;
using System.Data;
using MySql.Data.MySqlClient;

public class Handler
{
    private string connectionString;

    public Handler(string connectionString)
    {
        this.connectionString = connectionString;
    }

    public void Create(string query)
    {
        using (MySqlConnection connection = new MySqlConnection(connectionString))
        {
            connection.Open();
            MySqlCommand command = new MySqlCommand(query, connection);
            command.ExecuteNonQuery();
        }
    }

    public DataTable Read(string query)
    {
        DataTable dataTable = new DataTable();
        using (MySqlConnection connection = new MySqlConnection(connectionString))
        {
            connection.Open();
            MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
            adapter.Fill(dataTable);
        }
        return dataTable;
    }

    public void Update(string query)
    {
        using (MySqlConnection connection = new MySqlConnection(connectionString))
        {
            connection.Open();
            MySqlCommand command = new MySqlCommand(query, connection);
            command.ExecuteNonQuery();
        }
    }

    public void Delete(string query)
    {
        using (MySqlConnection connection = new MySqlConnection(connectionString))
        {
            connection.Open();
            MySqlCommand command = new MySqlCommand(query, connection);
            command.ExecuteNonQuery();
        }
    }

    public int Execute(string query)
    {
        int rowsAffected = 0;
        using (MySqlConnection connection = new MySqlConnection(connectionString))
        {
            connection.Open();
            MySqlCommand command = new MySqlCommand(query, connection);
            rowsAffected = command.ExecuteNonQuery();
        }
        return rowsAffected;
    }
}