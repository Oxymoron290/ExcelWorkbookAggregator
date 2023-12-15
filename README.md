# Excel Workbook Aggregator

A C# utility application for aggregating multiple input excel workbooks or csv files into one output workbook.

This application can take in either excel files or csv files and output either. The file processor is driven based on the file extensions.

## Usage

```
dotnet ExcelWorkbookAggregator.dll <options>

Options:
  -i, --input      Required. Input files to be processed.

  -o, --output     Required. File to output the result to.

  -l, --verbose    Set output to verbose messages.

  -h, --header     Indicates weather the worksheets include a header row.

  --help           Display this help screen.

  --version        Display version information.
```

Example #1:

```
.\src> dotnet run --l --i E:\team_blue.xlsx E:\team_red.xlsx "this is my file.xlsx" --o output.xlsx
```

Example #2:

```
.\src> dotnet run --l --i E:\team_blue_Teams.csv "E:\team_blue_Team Members.csv" E:\team_red_Teams.csv "E:\team_red_Team Members.csv" --o output.csv
```

