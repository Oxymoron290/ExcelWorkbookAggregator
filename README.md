# Excel Workbook Aggregator

A C# utility application for aggregating multiple input excel workbooks into one output workbook.

## Usage

```
dotnet ExcelWorkbookAggregator.dll <options>

Options:
  -i, --input      Required. Input files to be processed.

  -o, --output     Required. File to output the result to.

  -l, --verbose    Set output to verbose messages.

  --help           Display this help screen.

  --version        Display version information.
```

Example: `dotnet run --l --i E:\team_blue.xlsx E:\team_red.xlsx "this is my file.xlsx" --o output.xlsx``