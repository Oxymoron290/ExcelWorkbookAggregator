using CsvHelper;
using CsvHelper.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using Microsoft.Extensions.Logging;

namespace ExcelWorkbookAggregator;
public class CsvFileProcessor : IFileProcessor
{
    private readonly ILogger<CsvFileProcessor> _logger;
    public CsvFileProcessor(ILogger<CsvFileProcessor> logger)
    {
        _logger = logger;
    }

    public IEnumerable<DataTable> ProcessFile(string filePath, bool header)
    {
        var worksheet = Path.GetFileNameWithoutExtension(filePath);
        worksheet = worksheet.Split('_').Last();
        _logger.LogInformation($"Processing {worksheet} started.");
        var dataTable = new DataTable(worksheet);
        using (var reader = new StreamReader(filePath))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            // Assuming the first record is the header
            if (header)
            {
                csv.Read();
                csv.ReadHeader();
                foreach (var h in csv.HeaderRecord)
                {
                    _logger.LogInformation($"{h} added");
                    dataTable.Columns.Add(h);
                }
            }

            while (csv.Read())
            {
                var row = dataTable.NewRow();
                foreach (DataColumn column in dataTable.Columns)
                {
                    row[column.ColumnName] = csv.GetField(column.DataType, column.ColumnName);
                }
                dataTable.Rows.Add(row);
            }
        }

        foreach (DataRow dataRow in dataTable.Rows)
        {
            var msg = "Data: ";
            foreach (var item in dataRow.ItemArray)
            {
                msg += $"[{item}]";
            }
            _logger.LogInformation(msg);
        }

        _logger.LogInformation($"Processing {worksheet} completed.");
        return new List<DataTable>{dataTable};
    }

    public void WriteToFile(IEnumerable<DataTable> data, string filePath)
    {
        // Extract directory and base file name
        string directory = Path.GetDirectoryName(filePath) ?? "./";
        string baseFileName = Path.GetFileNameWithoutExtension(filePath);
        _logger.LogInformation($"Writing output to {directory} started.");

        foreach (DataTable table in data)
        {
            string newFileName = Path.Combine(directory, $"{baseFileName}_{table.TableName}.csv");
            _logger.LogInformation($"Creating worksheet {newFileName}");

            using (var writer = new StreamWriter(newFileName))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                // Write the headers
                foreach (DataColumn column in table.Columns)
                {
                    csv.WriteField(column.ColumnName);
                }
                csv.NextRecord();

                // Write the data
                foreach (DataRow row in table.Rows)
                {
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        csv.WriteField(row[i]);
                    }
                    csv.NextRecord();
                }
            }
            
            _logger.LogInformation($"Worksheet {newFileName} created!");
        }
    }
}
