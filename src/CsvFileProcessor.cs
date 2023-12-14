using System.Data;
using Microsoft.Extensions.Logging;

namespace ExcelWorkbookAggregator;
public class CsvFileProcessor : IFileProcessor
{
    private readonly ILogger<CsvFileProcessor> _logger;
    public CsvFileProcessor(ILogger<CsvFileProcessor> logger)
    {
        _logger = logger;
    }

    public DataTable ProcessFile(string filePath)
    {
        throw new NotImplementedException();
    }

    public void WriteToFile(DataTable data, string filePath)
    {
        throw new NotImplementedException();
    }
}
