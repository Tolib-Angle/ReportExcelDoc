using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using Izdel;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        List<OutputData> outputData = new List<OutputData>();
        // summ cost
        decimal cost = 0;
        // connection string insert into [ your_connection_string ]
        string connectionString = @"your_connection_string";
        // sql query get id first LVL
        string izdelQuery = "SELECT * FROM Izdel WHERE Id NOT IN (SELECT Izdel FROM Links)";

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        try
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Отчет");

                // Заголовки столбцов
                worksheet.Cells[1, 1].Value = "Изделие";
                worksheet.Cells[1, 2].Value = "Количество";
                worksheet.Cells[1, 3].Value = "Стоимость";
                worksheet.Cells[1, 4].Value = "Цена";

                worksheet.Column(1).Width = 30;
                worksheet.Column(2).Width = 15;
                worksheet.Column(3).Width = 15;
                worksheet.Column(4).Width = 15;

                int currentRow = 2;

                // connect DB and open connect
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand izdelCommand = new SqlCommand(izdelQuery, connection))
                    {
                        // Read Id from Izdel LVL 1
                        using (SqlDataReader izdelReader = izdelCommand.ExecuteReader())
                        {
                            while (izdelReader.Read())
                            {
                                OutputData _outputData = new OutputData();

                                _outputData.Id = izdelReader.GetInt64(0);
                                _outputData.Name = izdelReader.GetString(1);
                                _outputData.Kol = 1;
                                _outputData.Price = izdelReader.GetDecimal(2);

                                outputData.Add(_outputData);
                            }

                            izdelReader.Close();
                        }
                    }
                    // Search LVL2, LVL3 ... for Izdel LVL 1
                    for (int i = 0; i < outputData.Count; i++)
                    {
                        int tmpRow = currentRow;

                        worksheet.Cells[currentRow, 1].Value = outputData[i].Name;
                        worksheet.Cells[currentRow, 2].Value = outputData[i].Kol;
                        worksheet.Cells[currentRow, 4].Value = outputData[i].Price;

                        currentRow++;

                        cost = GetCosts(outputData[i].Id, connection, 1, ref currentRow, worksheet);

                        worksheet.Cells[tmpRow, 3].Value = cost;
                    }

                    package.SaveAs(new System.IO.FileInfo("Отчет.xlsx"));
                    Console.WriteLine("Отчет успешно сформирован!");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Произошла ошибка: {ex.Message}");
        }

        Console.ReadLine();
    }

    static decimal GetCosts(long id, SqlConnection connection, int lvl, ref int currentRow, ExcelWorksheet worksheet)
    {
        if (lvl >= 3)
            return 0;

        decimal costs = 0;
        List<OutputData> outputData = new List<OutputData> ();

        string izdelQuery = $"select Id, Name, kol, Price from Izdel inner join Links on Izdel.id = Links.Izdel where IzdelUp = {id}";

        try
        {
            using (SqlCommand izdelCommand = new SqlCommand(izdelQuery, connection))
            {
                using (SqlDataReader izdelReader = izdelCommand.ExecuteReader())
                {
                    while (izdelReader.Read())
                    {
                        OutputData _outputData = new OutputData();

                        _outputData.Id = izdelReader.GetInt64(0);
                        _outputData.Name = izdelReader.GetString(1);
                        _outputData.Kol = izdelReader.GetInt32(2);
                        _outputData.Price = izdelReader.GetDecimal(3);
                        _outputData.Cost = _outputData.Kol * _outputData.Price;

                        outputData.Add(_outputData);

                    }

                    izdelReader.Close();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Произошла ошибка: {ex.Message}");
        }

        for(int i = 0; i < outputData.Count; i++)
        {
            int tmpRow = currentRow;
            string tabs = "";

            for (int j = 0; j < lvl; j++)
                tabs += "   ";

            worksheet.Cells[currentRow, 1].Value = tabs + outputData[i].Name;
            worksheet.Cells[currentRow, 2].Value = outputData[i].Kol;
            worksheet.Cells[currentRow, 4].Value = outputData[i].Price;

            currentRow++;

            outputData[i].Cost += GetCosts(outputData[i].Id, connection, lvl + 1, ref currentRow, worksheet);

            worksheet.Cells[tmpRow, 3].Value = outputData[i].Cost;

            costs += outputData[i].Cost;
        }

        return costs;
    }
}