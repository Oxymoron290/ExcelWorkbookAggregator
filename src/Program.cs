using System;
using System.Reflection;
using System.Data;
using CommandLine;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Serilog;
namespace ExcelWorkbookAggregator;

class Program
{
    static void Main(string[] args)
    {
        Parser.Default.ParseArguments<Options>(args)
            .WithParsed<Options>(opts =>
            {
                if (opts.Version)
                {
                    var version = Assembly.GetEntryAssembly()?.GetName().Version;
                    Console.WriteLine(version?.ToString());
                    return;
                }
                var serviceProvider = ConfigureServices(args, opts);
                var app = serviceProvider.GetService<App>();
                app?.RunWithOptions(opts, serviceProvider);
            });
    }

    private static IServiceProvider ConfigureServices(string[] args, Options opts)
    {
        IConfigurationBuilder configBuilder = new ConfigurationBuilder()
            .SetBasePath(Path.Combine(AppContext.BaseDirectory))
            .AddJsonFile("appsettings.json", true)
            .AddJsonFile("appsettings.dev.json", true)
            .AddEnvironmentVariables()
            .AddCommandLine(args);

        IConfiguration config = configBuilder.Build();

        var loggerConfig = new LoggerConfiguration()
            .WriteTo.File("log.txt");
            
        if (opts.Verbose)
        {
            loggerConfig.WriteTo.Console();
            Console.WriteLine("The application is in Verbose mode!");
        }
        Log.Logger = loggerConfig.CreateLogger();
        var services = new ServiceCollection();
        services.AddSingleton(config);
        services.AddLogging(builder => builder.AddSerilog(dispose: true));

        services.AddTransient<App>();
        services.AddTransient<CsvFileProcessor>();
        services.AddTransient<ExcelFileProcessor>();

        return services.BuildServiceProvider();
    }
}
