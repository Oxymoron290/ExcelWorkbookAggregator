using System;
using CommandLine;
namespace ExcelWorkbookAggregator;

public class Options
{
    [Option('i', "input", Required = true, HelpText = "Input files to be processed (CSV or Excel).")]
    public IEnumerable<string> InputFiles { get; set; }

    [Option('h', "header", Required = false, Default = true, HelpText = "Indicates weather the worksheets include a header row.")]
    public bool Header { get; set; }

    [Option('o', "output", Required = true, HelpText = "File to output the result to (CSV or Excel).")]
    public string OutputFile { get; set; }

    [Option('l', "verbose", HelpText = "Enable verbose output.")]
    public bool Verbose { get; set; }

    [Option('v', "version", HelpText = "Print out the version of the program.")]
    public bool Version { get; set; }
}