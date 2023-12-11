using System;
using CommandLine;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Serilog;
namespace ExcelWorkbookAggregator;

class Program
{
    static void Main(string[] args)
    {
        var serviceProvider = ConfigureServices(args);
    }

    private static IServiceProvider ConfigureServices(string[] args)
    {
        Parser.Default.ParseArguments<Options>(args)
            .WithParsed<Options>(o =>
            {
                if (o.Verbose)
                {
                    Console.WriteLine("The application is in Verbose mode!");
                }
                foreach (var file in o.InputFiles)
                {
                    Console.WriteLine("Processing file: " + file);
                }
                Console.WriteLine("Hello, World!");
            });
            
        IConfigurationBuilder configBuilder = new ConfigurationBuilder()
            .SetBasePath(Path.Combine(AppContext.BaseDirectory))
            .AddJsonFile("appsettings.json", true)
            .AddJsonFile("appsettings.dev.json", true)
            .AddEnvironmentVariables()
            .AddCommandLine(args);

        IConfiguration config = configBuilder.Build();

        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console()
            .WriteTo.File("log.txt")
            .CreateLogger();
        var services = new ServiceCollection();
        services.AddSingleton(config);
        services.AddLogging(builder => builder.AddSerilog(dispose: true));

        //services.AddTransient<CustomDependencyClass>();

        return services.BuildServiceProvider();
    }
}
