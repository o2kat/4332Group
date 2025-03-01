using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

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

        // Кнопка для импорта данных из Excel в SQL
        private void ImportDataButton_Click(object sender, RoutedEventArgs e)
        {
            // Путь к файлу Excel
            string filePath = @"C:\Path\To\1.xlsx";  // Замените на путь к вашему файлу

            var excelApp = new Excel.Application();
            excelApp.Visible = false; // Окно Excel не показываем
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            // Подключение к базе данных SQL
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                for (int row = 2; row <= range.Rows.Count; row++) // Начинаем с 2 строки (чтобы пропустить заголовки)
                {
                    string serviceName = range.Cells[row, 1].Value.ToString();
                    decimal cost = Convert.ToDecimal(range.Cells[row, 2].Value);

                    // Вставка данных в таблицу базы данных
                    string query = "INSERT INTO ImportedData (ServiceName, Cost) VALUES (@ServiceName, @Cost)";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@ServiceName", serviceName);
                        cmd.Parameters.AddWithValue("@Cost", cost);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            // Закрываем Excel
            workbook.Close(false);
            excelApp.Quit();

            MessageBox.Show("Данные успешно импортированы в базу данных!");
        }

        // Обработчик кнопки для экспорта данных
        private void ExportDataButton_Click(object sender, RoutedEventArgs e)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true; // Показываем окно Excel
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
    }

    // Класс для хранения данных услуги
    public class ServiceData
    {
        public string ServiceName { get; set; }
        public decimal Cost { get; set; }
    }
}
