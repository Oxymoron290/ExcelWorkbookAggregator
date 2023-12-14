using System.Data;

namespace ExcelWorkbookAggregator;
public interface IFileProcessor
{
    DataTable ProcessFile(string filePath);
    void WriteToFile(DataTable data, string filePath);
}
