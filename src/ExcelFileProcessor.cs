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

    public IEnumerable<DataTable> ProcessFile(string filePath, bool header)
    {
        _logger.LogInformation($"Processing file: {filePath}");
        var result = new List<DataTable>();

        using (var workbook = new XLWorkbook(filePath))
            foreach (IXLWorksheet worksheet in workbook.Worksheets)
            {
                DataTable wsDataTable = ProcessWorksheet(worksheet, header);
                result.Add(wsDataTable);
            }
        _logger.LogInformation($"Processing {filePath} completed");

        return result;
    }

    private DataTable ProcessWorksheet(IXLWorksheet worksheet, bool header)
    {
        _logger.LogInformation($"Processing {worksheet.Name} started.");
        var dataTable = new DataTable(worksheet.Name);

        // Determine the number of columns
        int numberOfColumns = worksheet.RowsUsed().Max(r => r.LastCellUsed()?.Address.ColumnNumber ?? 1);

        _logger.LogInformation($"{numberOfColumns} columns");

        // Add columns
        bool columnsAdded = false;
        foreach (IXLRow row in worksheet.RowsUsed())
        {
            if (!columnsAdded)
            {
                columnsAdded = true;
                if (header)
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        var dc = new DataColumn(cell.Value.ToString());
                        dataTable.Columns.Add(dc); // Using column letters as column names
                    }
                    continue;
                }
                else
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        dataTable.Columns.Add(cell.Address.ColumnLetter); // Using column letters as column names
                    }
                }
            }

            // Add row data
            var dataRow = dataTable.NewRow();
            var i = 0;
            foreach (var cell in row.Cells(1, numberOfColumns))
            {
                dataRow[i++] = (!cell.Value.IsBlank) ? cell.Value : "";
            }
            // dataRow.ItemArray = row.Cells(1, numberOfColumns).Select(cell => ((!cell.Value.IsBlank) ? cell.Value : "").ToString()).ToArray();
            dataTable.Rows.Add(dataRow);
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
                    var cellValue = table.Rows[i][j];
                    worksheet.Cell(i + 2, j + 1).Value = cellValue?.ToString() ?? "";
                }
            }

            // Optional: Adjust column widths to content
            worksheet.Columns().AdjustToContents();
        }

        workbook.SaveAs(filePath);
        _logger.LogInformation($"{filePath} successfully created!");
    }
}