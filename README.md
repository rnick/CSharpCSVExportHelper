# CSharpCSVExportHelper
CSharp-Class to export a List&lt;> to CSV

## Example

```
// convert List to csv
List<CSVItem> data = getCSVData();

String CSVFileContent = CSVExportHelper.Generate(data.Cast<object>().ToList(), true, ",", "\"", true);
```
