using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;  // Explicitly using System.Data
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace GUI
{
    public partial class notary_record : Form
    {
        private Handler dbHandler;
        private string selectedTemplateFilePath = "";

        public notary_record()
        {
            // Set the license context for non-commercial use using the fully qualified name
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            InitializeComponent();
            dbHandler = new Handler("server=127.0.0.1; uid=root; pwd=Sparrowcakes212024; database=lawoffice");
            LoadDataFromDatabase();
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

        private void LoadDataFromDatabase()
        {
            try
            {
                string query = "SELECT recordID, entryID, file_title, total_file FROM filerecord";
                System.Data.DataTable dataTable = dbHandler.Read(query);
                dataGrid_Nrecord.DataSource = dataTable;
                PlotGraph(dataTable); // Call to plot graph based on data table
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading data: " + ex.Message);
            }

            foreach (DataGridViewColumn column in dataGrid_Nrecord.Columns)
            {
                column.Width = 124; // Set column width
            }
        }

        private void dataGrid_Nrecord_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Placeholder for functionality when a cell is clicked
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

                templateWorksheet = templateWorkbook.Sheets["Files Report"];
                sheet2 = templateWorkbook.Sheets["Sheet2"];

                if (templateWorksheet == null || sheet2 == null)
                {
                    MessageBox.Show("One or more required sheets were not found in the template workbook.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int headerRow = 6;
                int startRow = 7;
                int recordIDColumn = 0, entryIDColumn = 0, fileTitleColumn = 0, totalFilesColumn = 0;

                for (int i = 1; i <= templateWorksheet.UsedRange.Columns.Count; i++)
                {
                    var cellValue = (templateWorksheet.Cells[headerRow, i] as Range).Value2?.ToString();
                    switch (cellValue)
                    {
                        case "Record ID": recordIDColumn = i; break;
                        case "Entry ID": entryIDColumn = i; break;
                        case "Title of File": fileTitleColumn = i; break;
                        case "Number of Files": totalFilesColumn = i; break;
                    }
                }

                if (recordIDColumn * entryIDColumn * fileTitleColumn * totalFilesColumn == 0)
                {
                    MessageBox.Show("One or more required columns could not be found in the template.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                templateWorksheet.Cells[5, "D"] = DateTime.Now.ToString("MM-dd-yyyy");

                Dictionary<string, double> totalFilesByTitle = new Dictionary<string, double>();

                int currentRow = startRow;
                double sumFiles = 0;

                foreach (DataGridViewRow row in dataGrid_Nrecord.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        templateWorksheet.Cells[currentRow, recordIDColumn] = row.Cells["recordID"].Value?.ToString() ?? "";
                        templateWorksheet.Cells[currentRow, entryIDColumn] = row.Cells["entryID"].Value?.ToString() ?? "";
                        string fileTitle = (row.Cells["file_title"].Value?.ToString() ?? "").Trim();
                        double totalFiles = Convert.ToDouble(row.Cells["total_file"].Value ?? 0);
                        sumFiles += totalFiles; 

                        if (totalFilesByTitle.ContainsKey(fileTitle))
                            totalFilesByTitle[fileTitle] += totalFiles;
                        else
                            totalFilesByTitle[fileTitle] = totalFiles;

                        templateWorksheet.Cells[currentRow, fileTitleColumn] = fileTitle;
                        templateWorksheet.Cells[currentRow, totalFilesColumn].NumberFormat = "0";
                        templateWorksheet.Cells[currentRow, totalFilesColumn].Value2 = totalFiles;

                        currentRow++;
                    }
                }

                // Set the total in Excel cell H13
                templateWorksheet.Cells[13, "H"] = sumFiles;

                int sheet2Row = 2;
                foreach (var pair in totalFilesByTitle)
                {
                    sheet2.Cells[sheet2Row, 1].Value = pair.Key;
                    sheet2.Cells[sheet2Row, 2].Value = pair.Value;
                    sheet2Row++;
                }

                var chartObjects = (Microsoft.Office.Interop.Excel.ChartObjects)sheet2.ChartObjects(Type.Missing);
                var chartObject = chartObjects.Add(200, 50, 500, 500);
                var chart = chartObject.Chart;

                chart.SetSourceData(sheet2.Range[sheet2.Cells[2, 1], sheet2.Cells[sheet2Row - 1, 2]]);
                chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Total Count of Files Notarized";

                // Get the series collection and change color to gold
                var seriesCollection = (Microsoft.Office.Interop.Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);
                var series = seriesCollection.Item(1); // Assumes there's only one series in the chart

                // Set the color of the series to gold
                series.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Gold);

                // Optionally, you can format the plot area and chart area background if needed
                chart.PlotArea.Interior.Color = ColorTranslator.ToOle(Color.White);
                chart.ChartArea.Interior.Color = ColorTranslator.ToOle(Color.LightGray);

                chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

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
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;
            DialogResult userClickedOK = openFileDialog1.ShowDialog();

            if (userClickedOK == DialogResult.OK)
            {
                selectedTemplateFilePath = openFileDialog1.FileName;
                txb_openFile.Text = selectedTemplateFilePath;
            }
        }

        private void Record_graph_Click(object sender, EventArgs e)
        {
            // Placeholder for functionality when Record Graph button is clicked
        }

        private void PlotGraph(System.Data.DataTable dataTable)
        {
            var summary = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (System.Data.DataRow row in dataTable.Rows)
            {
                string title = row["file_title"] == DBNull.Value ? "Unknown" : row["file_title"].ToString().Trim();
                title = System.Text.RegularExpressions.Regex.Replace(title, @"\s+", " ");

                int count = row["total_file"] == DBNull.Value ? 0 : Convert.ToInt32(row["total_file"]);

                // Debugging output to check the processing of file titles and counts
                Console.WriteLine($"Processing: {title} with count {count}");

                if (summary.ContainsKey(title))
                {
                    summary[title] += count;
                }
                else
                {
                    summary[title] = count;
                }
            }

            Record_graph.Series.Clear();
            var series = new System.Windows.Forms.DataVisualization.Charting.Series("File Totals")
            {
                ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column
            };
            Record_graph.Series.Add(series);

            // Populate the series with aggregated data
            foreach (var pair in summary)
            {
                series.Points.AddXY(pair.Key, pair.Value);
                // Debugging output to verify the data added to the graph
                Console.WriteLine($"Adding to graph: {pair.Key} - {pair.Value}");
            }

            Record_graph.ChartAreas[0].AxisX.Interval = 1;
            Record_graph.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            Record_graph.ChartAreas[0].AxisX.LabelStyle.IsStaggered = false;
            Record_graph.ChartAreas[0].AxisX.LabelStyle.Enabled = true;
            Record_graph.ChartAreas[0].AxisY.Minimum = 0;

            Record_graph.Refresh();
        }





        private void btn_dashboard_Click(object sender, EventArgs e)
        {

        }

        private void btn_register_Click(object sender, EventArgs e)
        {

        }

        private void btn_book_Click(object sender, EventArgs e)
        {

        }

        private void btn_notaryFee_Click(object sender, EventArgs e)
        {

        }

        private void btn_reports_Click(object sender, EventArgs e)
        {

        }

        private void btn_about_Click(object sender, EventArgs e)
        {

        }
    }
}
