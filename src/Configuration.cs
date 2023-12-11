using Microsoft.Extensions.Configuration;
namespace ExcelWorkbookAggregator;
public class Configuration
{
    public string SourceWorkbookTab { get; private set; }
    public string SourceWorkbookColumn { get; private set; }
    public string DestinationWorkbookTab { get; private set; }
    public string DestinationWorkbookColumn { get; private set; }

    public Configuration(IConfiguration config)
    {
        SourceWorkbookTab = config["RecordMap:SourceWorkbookTab"] ?? "";
        SourceWorkbookColumn = config["RecordMap:SourceWorkbookColumn"] ?? "";
        DestinationWorkbookTab = config["RecordMap:DestinationWorkbookTab"] ?? "";
        DestinationWorkbookColumn = config["RecordMap:DestinationWorkbookColumn"] ?? "";
    }
}