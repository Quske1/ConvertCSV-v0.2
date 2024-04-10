using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ConvertCSV_v0._2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string inputFolderPath = folderBrowserDialog.SelectedPath;
                var csvFiles = Directory.GetFiles(inputFolderPath, "postings*.csv");
                int fileCounter = 1;
                foreach (var inputFile in csvFiles)
                {
                    // Формируем имя выходного файла
                    var outputFile = Path.Combine(inputFolderPath, $"Лист подбора OZON ({fileCounter}).xlsx");

                    // Номера столбцов, которые нужно удалить.
                    int[] columnsToRemove = { 0, 2, 3, 4, 5, 6, 7, 8, 10, 12, 13, 14, 15, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27 };

                    // Чтение данных из файла.
                    var fileLines = File.ReadAllLines(inputFile);

                    // Получаем все строки данных.
                    var records = fileLines.Skip(1) // пропуск строки заголовка
                    .Select(line => line.Split(';')) // разделение строк на массивы строк
                    .ToArray();

                    // Удаляем нужные столбцы.
                    var filteredRecords = records.Select(row =>
                    row.Where((value, index) => !columnsToRemove.Contains(index)).ToArray()
                    );

                    // Преобразование данных для записи в файл.
                    var transformedRecords = filteredRecords
                    .Where(record => record.Length > 0)
                    .ToList();

                    // Выводим информацию о количестве строк данных.
                    // Console.WriteLine($"Total rows: {transformedRecords.Count}");

                    // Выводим преобразованные строки в консоль.
                    // foreach (var row in transformedRecords)
                    //{
                    // Console.WriteLine(string.Join(",", row));
                    ///}

                    // Создаём новый документ Excel
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(new FileInfo(outputFile)))
                    {
                        var worksheet = package.Workbook.Worksheets.Add("Data");

                        // Вставляем заголовки в таблицу
                        var headerRow = new string[] { "Номер отправления", "Наименование товара", "Артикул", "Кол-во" };
                        for (int i = 0; i < headerRow.Length; i++)
                        {
                            worksheet.Cells[1, i + 1].Value = headerRow[i];
                        }

                        // Вставляем данные в таблицу
                        for (int i = 0; i < transformedRecords.Count; i++)
                        {
                            var row = transformedRecords[i];
                            for (int j = 0; j < row.Length; j++)
                            {
                                worksheet.Cells[i + 2, j + 1].Value = row[j];
                            }
                        }

                        // Автоподгоняем ширину ячеек под содержимое
                        worksheet.Cells.AutoFitColumns();

                        for (int i = 2; i <= transformedRecords.Count + 1; i++)
                        {
                            for (int j = 1; j <= 4; j++)
                            {
                                if (worksheet.Cells[i, j].Value != null)
                                {
                                    string value = worksheet.Cells[i, j].Value.ToString();
                                    value = value.Replace("\"", "");
                                    worksheet.Cells[i, j].Value = value;
                                }
                            }
                        }
                        // Выравниваем содержимое в ячейках
                        var dataRange = worksheet.Cells[2, 1, transformedRecords.Count + 1, 4];
                        dataRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        for (int i = 2; i <= transformedRecords.Count + 1; i++)
                        {
                            string value = worksheet.Cells[i, 1].Value.ToString();
                            int startRow = i;
                            int endRow = i;
                            for (int j = i + 1; j <= transformedRecords.Count + 1; j++)
                            {
                                if (worksheet.Cells[j, 1].Value.ToString() == value)
                                {
                                    endRow = j;
                                }
                                else
                                {
                                    break;
                                }
                            }
                            if (startRow != endRow)
                            {
                                var mergeRange = worksheet.Cells[startRow, 1, endRow, 1];
                                mergeRange.Merge = true;
                            }
                            i = endRow;
                        }

                        package.Save();

                        // Сохраняем изменения в документе Excel
                        var dataCells = worksheet.Cells[1, 1, transformedRecords.Count + 1, 4];
                        dataCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        dataCells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        dataCells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        dataCells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        // Рисуем границы вокруг заголовков таблицы
                        var headerCells = worksheet.Cells[1, 1, 1, 4];
                        headerCells.Style.Font.Bold = true;
                        headerCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        headerCells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(218, 218, 218));
                        headerCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        headerCells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        headerCells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        headerCells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Column(1).Width = 19;
                        worksheet.Column(2).Width = 37;
                        worksheet.Column(3).Width = 22;
                        worksheet.Column(4).Width = 7;
                        var wrapColumns = new int[] { 2 };
                        foreach (var column in wrapColumns)
                        {
                            var columnRange = worksheet.Cells[2, column, transformedRecords.Count + 1, column];
                            columnRange.Style.WrapText = true;
                        }

                        dataCells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        string[] boldStrings = { "Набор Ароматизаторов для автомобиля 2 шт", "Ароматизаторы в машину 3 шт", "2 шт по 600 мл", "насекомых и клещей 3 шт", "комплект 3 шт.", "Ароматизаторов в машину 3 шт" };
                        for (int i = 2; i <= transformedRecords.Count + 1; i++)
                        {
                            for (int j = 1; j <= 4; j++)
                            {
                                if (worksheet.Cells[i, j].Value != null)
                                {
                                    string value = worksheet.Cells[i, j].Value.ToString();
                                    value = value.Replace("\"", "");
                                    worksheet.Cells[i, j].Value = value;
                                    if (boldStrings.Any(s => value.Contains(s)))
                                    {
                                        worksheet.Cells[i, j].Style.Font.Bold = true; // добавляем жирный стиль
                                    }
                                }
                            }
                        }
                        package.SaveAs(new FileInfo(outputFile));
                        // Console.WriteLine($"File {outputFile} created successfully");

                        fileCounter++;
                    }
                }

                // Console.WriteLine("Done!");
                // Console.ReadKey();
                System.Diagnostics.Process.GetCurrentProcess().StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                Environment.Exit(0);

            }

        }
    }
}

