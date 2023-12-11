using System;
using CommandLine;
namespace ExcelWorkbookAggregator;

public class Options
{
    [Option('i', "input", Required = true, HelpText = "Input files to be processed.")]
    public IEnumerable<string> InputFiles { get; set; }

    [Option('o', "output", Required = true, HelpText = "File to output the result to.")]
    public string OutputFile { get; set; }

    [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
    public bool Verbose { get; set; }
}