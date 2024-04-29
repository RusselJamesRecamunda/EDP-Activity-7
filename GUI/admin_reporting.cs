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
    public partial class admin_reporting : Form
    {
        private Handler dbHandler;
        private string selectedTemplateFilePath = "";
        public admin_reporting()
        {
            InitializeComponent();
            // Initialize the database handler with your connection string
            string connectionString = "server=127.0.0.1; uid=root; pwd=Sparrowcakes212024; database=lawoffice";
            dbHandler = new Handler(connectionString);

            // Load data into DataGridView when the form loads
            LoadDataIntoDataGridView();
        }

        private void btn_Nregister_Click(object sender, EventArgs e)
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

        private void txb_file_TextChanged(object sender, EventArgs e)
        {

        }

        private void txb_Nparties_TextChanged(object sender, EventArgs e)
        {

        }

        private void txb_Aparties_TextChanged(object sender, EventArgs e)
        {

        }

        private void txb_id_TextChanged(object sender, EventArgs e)
        {

        }

        private void btn_add_notary_Click(object sender, EventArgs e)
        {
            try
            {
                string fileTitle = txb_file.Text;
                string notaryFee = txb_fee.Text;
                string totalFiles = txb_total.Text;
                string nameOfParties = txb_Nparties.Text;
                string addressOfParties = txb_Aparties.Text;
                string idNumbers = txb_id.Text;

                // Check if any of the textboxes are empty
                if (string.IsNullOrEmpty(fileTitle) || string.IsNullOrEmpty(notaryFee) || string.IsNullOrEmpty(totalFiles) || 
                    string.IsNullOrEmpty(nameOfParties) || string.IsNullOrEmpty(addressOfParties) || string.IsNullOrEmpty(idNumbers))
                {
                    MessageBox.Show("Please fill in all the fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return; // Exit the method if any textbox is empty
                }

                DateTime notaryDate = DateTime.Now; 
                TimeSpan notaryTime = DateTime.Now.TimeOfDay; 

                string query = "INSERT INTO notarial_register (file_title, notary_fee, total_files, name_of_parties, address_of_parties, ID_numbers, notary_date, notary_time) " +
                               "VALUES ('" + fileTitle + "', '" + notaryFee + "', '" + totalFiles + "',  '" + nameOfParties + "', '" + addressOfParties + "', '" + idNumbers + "', '" + notaryDate.ToString("yyyy-MM-dd") + "', '" + notaryTime.ToString() + "')";

                dbHandler.Execute(query);

                MessageBox.Show("Notary register successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadDataIntoDataGridView()
        {
            string query = "SELECT entryID, file_title, notary_fee, total_files, name_of_parties, address_of_parties, ID_numbers, notary_date, notary_time FROM notarial_register";
            System.Data.DataTable data = dbHandler.Read(query);
            dataGrid_Nregister.DataSource = data;

            // Set the width of each column to 140 pixels
            foreach (DataGridViewColumn column in dataGrid_Nregister.Columns)
            {
                column.Width = 131 ;
            }
        }

        private void dataGrid_Nregister_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void btn_reports_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

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
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;
                templateWorkbook = excelApp.Workbooks.Open(selectedTemplateFilePath);
                templateWorksheet = templateWorkbook.Sheets["Financial Notary Report"];
                sheet2 = templateWorkbook.Sheets["Sheet2"];

                if (templateWorksheet == null || sheet2 == null)
                {
                    MessageBox.Show("One or more required sheets were not found in the template workbook.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int headerRow = 6;
                int startRow = 7;
                int entryIDColumn = 0, fileTitleColumn = 0, notaryFeeColumn = 0, nameOfPartiesColumn = 0, addressOfPartiesColumn = 0, notaryDateColumn = 0, notaryTimeColumn = 0;

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
                        case "Notary Fee":
                            notaryFeeColumn = i;
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
                        case "Time of Notary":
                            notaryTimeColumn = i;
                            break;
                    }
                }

                if (entryIDColumn * fileTitleColumn * notaryFeeColumn * nameOfPartiesColumn * addressOfPartiesColumn * notaryDateColumn * notaryTimeColumn == 0)
                {
                    MessageBox.Show("One or more required columns could not be found in the template.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                templateWorksheet.Cells[5, "D"] = DateTime.Now.ToString("MM-dd-yyyy");

                Dictionary<string, double> notaryFeesByTitle = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                double totalNotaryFee = 0;  // Initialize total fees

                int currentRow = startRow;
                foreach (DataGridViewRow row in dataGrid_Nregister.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string fileTitle = (row.Cells["file_title"].Value?.ToString() ?? "").Trim();

                        string notaryFeeWithPeso = "₱" + (row.Cells["notary_fee"].Value?.ToString() ?? "");
                        templateWorksheet.Cells[currentRow, notaryFeeColumn] = notaryFeeWithPeso;

                        if (double.TryParse(row.Cells["notary_fee"].Value?.ToString(), out double fee))
                        {
                            if (notaryFeesByTitle.ContainsKey(fileTitle))
                                notaryFeesByTitle[fileTitle] += fee;
                            else
                                notaryFeesByTitle[fileTitle] = fee;

                            totalNotaryFee += fee;  // Add fee to the total
                        }


                        // Copying remaining cell values as per original code
                        templateWorksheet.Cells[currentRow, entryIDColumn] = row.Cells["entryID"].Value?.ToString() ?? "";
                        templateWorksheet.Cells[currentRow, fileTitleColumn] = row.Cells["file_title"].Value?.ToString() ?? "";
                        templateWorksheet.Cells[currentRow, nameOfPartiesColumn] = row.Cells["name_of_parties"].Value?.ToString() ?? "";
                        templateWorksheet.Cells[currentRow, addressOfPartiesColumn] = row.Cells["address_of_parties"].Value?.ToString() ?? "";
                        templateWorksheet.Cells[currentRow, notaryDateColumn] = DateTime.TryParse(row.Cells["notary_date"].Value?.ToString(), out DateTime notaryDate) ? notaryDate.ToString("MM-dd-yyyy") : "";
                        templateWorksheet.Cells[currentRow, notaryTimeColumn] = DateTime.TryParse(row.Cells["notary_time"].Value?.ToString(), out DateTime notaryTime) ? notaryTime.ToString("hh:mm tt") : "";
                        currentRow++;
                    }
                }

                // Write total fee to cell K13
                templateWorksheet.Cells[13, "K"] = "₱" + totalNotaryFee.ToString("N2");

                int sheet2Row = 2;
                foreach (var pair in notaryFeesByTitle)
                {
                    sheet2.Cells[sheet2Row, 1].Value = pair.Key;
                    sheet2.Cells[sheet2Row, 2].Value = pair.Value;
                    sheet2Row++;
                }

                var chartObjects = (ChartObjects)sheet2.ChartObjects();
                var chartObject = chartObjects.Add(200, 50, 500, 500);
                var chart = chartObject.Chart;
                chart.SetSourceData(sheet2.Range[sheet2.Cells[2, 1], sheet2.Cells[sheet2Row - 1, 2]]);
                chart.ChartType = XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Total Fee Notarized Per Document";
                SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection();
                Series series = seriesCollection.Item(1);
                series.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Gold);

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

        private void txb_filedialog_TextChanged(object sender, EventArgs e)
        {

        }

        private void btn_register_Click(object sender, EventArgs e)
        {

        }
    }
}
