using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;
using System.IO;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace _4332Project
{
    public partial class MainWindow : Window
    {
        private string connectionString = "Server=O2KAT;Database=ISRPO;Trusted_Connection=True;";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void MukhametshinButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Имя: Азат\nВозраст: 18 лет\nГруппа: 4332\nКурс: 3\nСпециальность: Информационные технологии\nФакультет: Факультет информационных технологий");
        }

        private void ImportDataButton_Click(object sender, RoutedEventArgs e)
        {
            string filePath = @"C:\Users\o2kat\OneDrive\3 курс\ИСРПО\2lab\4332Group\4332Project\1.xlsx";

            var excelApp = new Excel.Application();
            excelApp.Visible = false;
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                for (int row = 2; row <= range.Rows.Count; row++)
                {
                    string serviceName = range.Cells[row, 1].Value.ToString();
                    decimal cost = Convert.ToDecimal(range.Cells[row, 2].Value);

                    string query = "INSERT INTO ImportedData (ServiceName, Cost) VALUES (@ServiceName, @Cost)";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@ServiceName", serviceName);
                        cmd.Parameters.AddWithValue("@Cost", cost);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            workbook.Close(false);
            excelApp.Quit();

            MessageBox.Show("Данные успешно импортированы в базу данных!");
        }

        private void ExportDataButton_Click(object sender, RoutedEventArgs e)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            List<ServiceData> serviceList = new List<ServiceData>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                string query = "SELECT ServiceName, Cost FROM ImportedData";
                using (SqlCommand cmd = new SqlCommand(query, conn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        serviceList.Add(new ServiceData
                        {
                            ServiceName = reader["ServiceName"].ToString(),
                            Cost = Convert.ToDecimal(reader["Cost"])
                        });
                    }
                }
            }

            var sortedList = serviceList.OrderBy(s => s.Cost).ToList();

            int row = 1;
            worksheet.Cells[row, 1].Value = "Название услуги";
            worksheet.Cells[row, 2].Value = "Стоимость";
            row++;

            foreach (var service in sortedList)
            {
                worksheet.Cells[row, 1].Value = service.ServiceName;
                worksheet.Cells[row, 2].Value = service.Cost;
                row++;
            }

            string filePath = @"C:\Users\o2kat\OneDrive\3 курс\ИСРПО\2lab\4332Group\4332Project\ExportedData.xlsx";
            workbook.SaveAs(filePath);
            workbook.Close(false);
            excelApp.Quit();

            MessageBox.Show($"Экспорт завершён! Файл сохранён в {filePath}");
        }

        private void ImportJsonDataButton_Click(object sender, RoutedEventArgs e)
        {
            string jsonFilePath = @"C:\Users\o2kat\OneDrive\3 курс\ИСРПО\2lab\4332Group\4332Project\1.json";
            string jsonData = File.ReadAllText(jsonFilePath);

            var services = JsonConvert.DeserializeObject<List<ServiceData>>(jsonData);

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                foreach (var service in services)
                {
                    string query = "INSERT INTO ImportedData (ServiceName, Cost) VALUES (@ServiceName, @Cost)";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@ServiceName", service.ServiceName);
                        cmd.Parameters.AddWithValue("@Cost", service.Cost);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            MessageBox.Show("Данные успешно импортированы из JSON в базу данных!");
        }

        private void ExportToWordButton_Click(object sender, RoutedEventArgs e)
        {
            List<ServiceData> serviceList = new List<ServiceData>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                string query = "SELECT ServiceName, Cost FROM ImportedData";
                using (SqlCommand cmd = new SqlCommand(query, conn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        serviceList.Add(new ServiceData
                        {
                            ServiceName = reader["ServiceName"].ToString(),
                            Cost = Convert.ToDecimal(reader["Cost"])
                        });
                    }
                }
            }

            var sortedList = serviceList.OrderBy(s => s.Cost).ToList();

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;

            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();

            Microsoft.Office.Interop.Word.Paragraph para = doc.Paragraphs.Add();
            para.Range.Text = "Отчет об услугах";
            para.Range.InsertParagraphAfter();

            decimal? lastCost = null;
            Microsoft.Office.Interop.Word.Table table = null;

            foreach (var service in sortedList)
            {
                if (lastCost != null && service.Cost != lastCost)
                {
                    para = doc.Paragraphs.Add();
                    para.Range.InsertParagraphAfter();
                    lastCost = service.Cost;
                }

                if (table == null || service.Cost != lastCost)
                {
                    lastCost = service.Cost;
                    para = doc.Paragraphs.Add();
                    para.Range.Text = $"Стоимость: {service.Cost}";

                    table = doc.Tables.Add(para.Range, 1, 2);
                    table.Borders.Enable = 1;
                    table.Rows[1].Cells[1].Range.Text = "Название услуги";
                    table.Rows[1].Cells[2].Range.Text = "Стоимость";
                }

                Microsoft.Office.Interop.Word.Row newRow = table.Rows.Add();
                newRow.Cells[1].Range.Text = service.ServiceName;
                newRow.Cells[2].Range.Text = service.Cost.ToString();
            }

            string wordFilePath = @"C:\Users\o2kat\OneDrive\3 курс\ИСРПО\2lab\4332Group\4332Project\ExportedData.docx";
            doc.SaveAs2(wordFilePath);
            doc.Close();
            wordApp.Quit();

            MessageBox.Show($"Экспорт завершён! Файл сохранён в {wordFilePath}");
        }

        public class ServiceData
        {
            public string ServiceName { get; set; }
            public decimal Cost { get; set; }
        }
    }
}
