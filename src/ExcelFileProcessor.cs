using System.Data;
using System.Xml;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging;

namespace ExcelWorkbookAggregator;
public class ExcelFileProcessor : IFileProcessor
{
    private readonly ILogger<ExcelFileProcessor> _logger;
    public ExcelFileProcessor(ILogger<ExcelFileProcessor> logger)
    {
        _logger = logger;
    }

    public IEnumerable<DataTable> ProcessFile(string filePath)
    {
        _logger.LogInformation($"Processing file: {filePath}");
        var result = new List<DataTable>();
        
        using (var workbook = new XLWorkbook(filePath))
        foreach (IXLWorksheet worksheet in workbook.Worksheets)
        {
            DataTable wsDataTable = ProcessWorksheet(worksheet);
            result.Add(wsDataTable);
        }
        _logger.LogInformation($"Processing {filePath} completed");

        return result;
    }
    
    private DataTable ProcessWorksheet(IXLWorksheet worksheet)
    {
        _logger.LogInformation($"Processing {worksheet.Name} started.");
        var dataTable = new DataTable(worksheet.Name);

        // Add columns
        bool columnsAdded = false;
        foreach (IXLRow row in worksheet.RowsUsed())
        {
            // var msg = "Data: ";
            // foreach (var cell in row.Cells())
            // {
            //     msg += $"[{cell}]{cell.Value} ";
            // }
            // _logger.LogInformation(msg);
            if (!columnsAdded)
            {
                foreach (IXLCell cell in row.Cells())
                {
                    dataTable.Columns.Add(cell.Address.ColumnLetter); // Using column letters as column names
                }
                columnsAdded = true;
            }

            // Add row data
            var dataRow = dataTable.NewRow();
            var i = 0;
            foreach (var cell in row.Cells()){
                dataRow[i++] = cell.Value;
            }
            dataTable.Rows.Add(dataRow);
            // dataTable.Rows.Add(row.Cells().Select(cell => cell.Value).ToArray());
        }

        foreach(DataRow dataRow in dataTable.Rows)
        {
            var msg = "Data: ";
            foreach(var item in dataRow.ItemArray)
            {
                msg += $"[{item}]";
            }
            _logger.LogInformation(msg);
        }

        _logger.LogInformation($"Processing {worksheet.Name} completed.");
        return dataTable;
    }

    public void WriteToFile(IEnumerable<DataTable> dataTables, string filePath)
    {
        _logger.LogInformation($"Writing output to {filePath} started.");
        using var workbook = new XLWorkbook();
        foreach (DataTable table in dataTables)
        {
            _logger.LogInformation($"Creating worksheet {table.TableName}");
            var worksheet = workbook.Worksheets.Add(table.TableName);

            // Adding headers
            for (int i = 0; i < table.Columns.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = table.Columns[i].ColumnName;
            }

            // Adding data
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    worksheet.Cell(i + 2, j + 1).Value = table.Rows[i][j].ToString();
                }
            }

            // Optional: Adjust column widths to content
            worksheet.Columns().AdjustToContents();
        }

        workbook.SaveAs(filePath);
        _logger.LogInformation($"{filePath} successfully created!");
    }
}