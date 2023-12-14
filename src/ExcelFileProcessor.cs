using System.Data;
using Microsoft.Extensions.Logging;

namespace ExcelWorkbookAggregator;
public class ExcelFileProcessor : IFileProcessor
{
    private readonly ILogger<ExcelFileProcessor> _logger;
    public ExcelFileProcessor(ILogger<ExcelFileProcessor> logger)
    {
        _logger = logger;
    }

    public DataTable ProcessFile(string filePath)
    {
        var result = new DataTable();
        // TODO: check that file exists and load the data
        // TODO: parse the data into a DataTable
        
        //throw new NotImplementedException();
        return result;
    }

    public void WriteToFile(DataTable data, string filePath)
    {
        // TODO: write the `data` to the `filePath`
        //throw new NotImplementedException();
        return;
    }
}