using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace GUI
{
    public partial class month_report : Form
    {
        private Handler databaseHandler;
        private string selectedTemplateFilePath = "";

        public month_report()
        {
            InitializeComponent();
            string connectionString = "server=127.0.0.1; uid=root; pwd=Sparrowcakes212024; database=lawoffice";
            databaseHandler = new Handler(connectionString);
            // Attach the DateSelected event handler
            monthCalendar1.DateSelected += MonthCalendar1_DateSelected;

            // Load data from the database
            LoadDataFromDatabase();
        }

        private void btn_openFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            DialogResult userClickedOK = openFileDialog1.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {
                // Get the selected file path and display it in the textbox.
                selectedTemplateFilePath = openFileDialog1.FileName;
                txb_filedialog.Text = selectedTemplateFilePath;
            }
        }

        private void btn_excel_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedTemplateFilePath))
            {
                MessageBox.Show("Please select a template Excel file first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Worksheet templateWorksheet = null;
            Worksheet sheet2 = null;
            Workbook templateWorkbook = null;
            Microsoft.Office.Interop.Excel.Application excelApp = null;

            try
            {
                // Create an Excel application instance
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;

                // Open the template workbook
                templateWorkbook = excelApp.Workbooks.Open(selectedTemplateFilePath);

                // Identify the "Month Report" sheet
                templateWorksheet = templateWorkbook.Sheets["Month Report"];
                sheet2 = templateWorkbook.Sheets["Sheet2"]; // Ensure that Sheet2 exists

                // Check if "Month Report" sheet exists in the template workbook
                if (templateWorksheet == null || sheet2 == null)
                {
                    MessageBox.Show("One or more required sheets were not found in the template workbook.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int headerRow = 6; // The row where your headers are located
                int startRow = 7; // The row where data starts
                int currentRow = startRow;
                int entryIDColumn = 0, fileTitleColumn = 0, totalFilesColumn = 0, nameOfPartiesColumn = 0, addressOfPartiesColumn = 0, notaryDateColumn = 0;

                // Identify the columns by header name
                for (int i = 1; i <= templateWorksheet.UsedRange.Columns.Count; i++)
                {
                    var cellValue = (templateWorksheet.Cells[headerRow, i] as Range).Value2?.ToString();
                    switch (cellValue)
                    {
                        case "Entry ID":
                            entryIDColumn = i;
                            break;
                        case "Title of File":
                            fileTitleColumn = i;
                            break;
                        case "Number of Files":
                            totalFilesColumn = i;
                            break;
                        case "Name of Parties":
                            nameOfPartiesColumn = i;
                            break;
                        case "Address of Parties":
                            addressOfPartiesColumn = i;
                            break;
                        case "Date of Notary":
                            notaryDateColumn = i;
                            break;
                    }
                }

                templateWorksheet.Cells[5, "D"] = DateTime.Now.ToString("MM-dd-yyyy");

                if (entryIDColumn * fileTitleColumn * totalFilesColumn * nameOfPartiesColumn * addressOfPartiesColumn * notaryDateColumn == 0)
                {
                    MessageBox.Show("One or more required columns could not be found in the template.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Dictionary to store total_files by file title
                Dictionary<string, double> totalFilesByTitle = new Dictionary<string, double>();

                // Write data from DataGridView to the template worksheet
                foreach (DataGridViewRow row in dataGrid_Rmonth.Rows)
                {
                    if (!row.IsNewRow) // Ensure the row is not the 'new' row
                    {
                        templateWorksheet.Cells[currentRow, entryIDColumn] = row.Cells["entryID"].Value?.ToString() ?? "";
                        string fileTitle = (row.Cells["file_title"].Value?.ToString() ?? "").Trim();
                        double totalFiles = Convert.ToDouble(row.Cells["total_files"].Value ?? 0);

                        templateWorksheet.Cells[currentRow, fileTitleColumn] = fileTitle;
                        templateWorksheet.Cells[currentRow, totalFilesColumn].NumberFormat = "0";
                        templateWorksheet.Cells[currentRow, totalFilesColumn].Value2 = totalFiles;

                        if (totalFilesByTitle.ContainsKey(fileTitle))
                            totalFilesByTitle[fileTitle] += totalFiles;
                        else
                            totalFilesByTitle[fileTitle] = totalFiles;

                        templateWorksheet.Cells[currentRow, nameOfPartiesColumn] = row.Cells["name_of_parties"].Value?.ToString() ?? "";
                        templateWorksheet.Cells[currentRow, addressOfPartiesColumn] = row.Cells["address_of_parties"].Value?.ToString() ?? "";

                        string notaryDateString = row.Cells["notary_date"].Value?.ToString();
                        templateWorksheet.Cells[currentRow, notaryDateColumn] = DateTime.TryParse(notaryDateString, out DateTime notaryDate) ? notaryDate.ToString("MM-dd-yyyy") : "";

                        currentRow++;
                    }
                }

                // Write aggregated data to Sheet2
                int sheet2Row = 2;
                foreach (var pair in totalFilesByTitle)
                {
                    sheet2.Cells[sheet2Row, 1].Value = pair.Key;
                    sheet2.Cells[sheet2Row, 2].Value = pair.Value;
                    sheet2Row++;
                }

                // Create and format the chart in Sheet2
                var chartObjects = (ChartObjects)sheet2.ChartObjects();
                var chartObject = chartObjects.Add(200, 50, 500, 500); // Position and size of chart
                var chart = chartObject.Chart;
                chart.SetSourceData(sheet2.Range[sheet2.Cells[2, 1], sheet2.Cells[sheet2Row - 1, 2]]);
                chart.ChartType = XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Total Files Notarized This Month";

                // Set the color of the series to gold
                SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection();
                Series series = seriesCollection.Item(1);
                series.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Gold);

                // Save and close
                string dateSuffix = DateTime.Now.ToString("MM-dd-yyyy");
                string newFileName = $"Notary_report_{dateSuffix}.xlsx";
                string savePath = Path.Combine(Path.GetDirectoryName(selectedTemplateFilePath), newFileName);
                templateWorkbook.SaveAs(savePath);

                MessageBox.Show("Data exported successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (templateWorksheet != null) ReleaseObject(templateWorksheet);
                if (sheet2 != null) ReleaseObject(sheet2);
                if (templateWorkbook != null)
                {
                    templateWorkbook.Close(false);
                    ReleaseObject(templateWorkbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    ReleaseObject(excelApp);
                }
            }
        }



        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                GC.Collect();
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            admin_reporting Areport = new admin_reporting();
            Areport.Show();
            this.Hide();
        }

        private void btn_Rmonth_Click(object sender, EventArgs e)
        {
            month_report Mreport = new month_report();
            Mreport.Show();
            this.Hide();
        }

        private void btn_Nrecord_Click(object sender, EventArgs e)
        {
            notary_record Nrecord = new notary_record();
            Nrecord.Show();
            this.Hide();
        }

        private void MonthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            // Extract the month and year from the selected date
            int month = monthCalendar1.SelectionStart.Month;
            int year = monthCalendar1.SelectionStart.Year;

            // Set the text of the textbox
            txb_Msearch.Text = $"{month}/{year}";
        }

        private void btn_search_Click(object sender, EventArgs e)
        {
            try
            {
                // Get the text from the search textbox
                string searchText = txb_Msearch.Text;

                // Split the text to extract month and year
                string[] parts = searchText.Split('/');
                if (parts.Length != 2)
                {
                    MessageBox.Show("Please enter a valid search term in the format 'MM/YYYY'.");
                    return;
                }

                if (!int.TryParse(parts[0], out int month) || !int.TryParse(parts[1], out int year))
                {
                    MessageBox.Show("Please enter a valid search term in the format 'MM/YYYY'.");
                    return;
                }

                // Construct the query to fetch data for the specified month and year
                string query = $"SELECT entryID, file_title, total_files, name_of_parties, address_of_parties, notary_date FROM " +
                    $"month_report WHERE MONTH(notary_date) = {month} AND YEAR(notary_date) = {year}";

                // Call Read method of databaseHandler to fetch data
                System.Data.DataTable data = databaseHandler.Read(query);

                // Bind data to the DataGridView
                dataGrid_Rmonth.DataSource = data;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadDataFromDatabase()
        {
            // Construct the query to fetch all data
            string query = "SELECT entryID, file_title, total_files, name_of_parties, address_of_parties, notary_date FROM month_report";

            // Call Read method of databaseHandler to fetch data
            System.Data.DataTable data = databaseHandler.Read(query);

            // Bind data to the DataGridView
            dataGrid_Rmonth.DataSource = data;

            // Set the width of each column to 140 pixels
            foreach (DataGridViewColumn column in dataGrid_Rmonth.Columns)
            {
                column.Width = 144;
            }
        }

        private void dataGrid_Rmonth_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void txb_filedialog_TextChanged(object sender, EventArgs e)
        {

        }

        private void txb_Msearch_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
