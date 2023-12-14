using System.Data;

namespace ExcelWorkbookAggregator;
public interface IFileProcessor
{
    IEnumerable<DataTable> ProcessFile(string filePath);
    void WriteToFile(IEnumerable<DataTable> data, string filePath);
}
