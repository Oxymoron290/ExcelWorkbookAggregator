using System.Data;

namespace ExcelWorkbookAggregator;
public interface IFileProcessor
{
    IEnumerable<DataTable> ProcessFile(string filePath, bool header);
    void WriteToFile(IEnumerable<DataTable> data, string filePath);
}
