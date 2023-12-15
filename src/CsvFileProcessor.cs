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

    public IEnumerable<DataTable> ProcessFile(string filePath, bool header)
    {
        var result = new List<DataTable>();
        // TODO: check that file exists and load the data
        // TODO: parse the data into a DataTable
        
        //throw new NotImplementedException();
        return result;
    }

    public void WriteToFile(IEnumerable<DataTable> data, string filePath)
    {
        throw new NotImplementedException();
    }
}
