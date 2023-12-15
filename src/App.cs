
using System.Data;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace ExcelWorkbookAggregator;

public class App
{
    private readonly ILogger<App> _logger;
    private readonly Configuration _config;
    public App(ILogger<App> logger, IConfiguration configuration)
    {
        _logger = logger;
        _config = new Configuration(configuration);
    }

    public void RunWithOptions(Options opts, IServiceProvider serviceProvider)
    {
        _logger.LogInformation("Processing Started!");
        var aggregatedData = new List<DataTable>();
        foreach (var file in opts.InputFiles)
        {
            var processor = getFileProcessor(serviceProvider, file);
            var fileData = processor?.ProcessFile(file);
            
            foreach(var table in fileData)
            {
                _logger.LogInformation($"Aggregating {table.TableName}");
                var targetTable = aggregatedData.FirstOrDefault(d => d.TableName == table.TableName);
                if(targetTable != null)
                {
                    targetTable.Merge(table);
                }
                else
                {
                    aggregatedData.Add(table);
                }
            }
        }

        var outputProcessor = getFileProcessor(serviceProvider, opts.OutputFile);
        outputProcessor?.WriteToFile(aggregatedData, opts.OutputFile);
        _logger.LogInformation("Processing Completed!");
    }
    
    private IFileProcessor? getFileProcessor(IServiceProvider serviceProvider, string filePath)
    {
        if(filePath.EndsWith(".xlsx"))
        {
            _logger.LogInformation($"ExcelFileProcessor selected for {filePath}");
            return serviceProvider.GetService<ExcelFileProcessor>();
        }
        else
        {
            _logger.LogInformation($"CsvFileProcessor selected for {filePath}");
            return serviceProvider.GetService<CsvFileProcessor>();
        }
    }
}