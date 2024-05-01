# ExcelSerializer
Convert object to Excel file (.xlsx) [Open XML SpreadsheetML File Format]

## Getting Started
Supporting platform is .NET Standard 2.0, 2.1, .NET 5, .NET 6., .NET 8.

## Usage
You can use `ExcelSerializer.ToFile` to create .xlsx file.

~~~
ExcelSerializer.ToFile(Users, "test.xlsx", ExcelSerializerOptions.Default);
~~~

## Notice

Folder creation permissions are required since a working folder will be used.

## Benchmark
N = 100 lines.

| Method              | N   | Mean        | Error      | StdDev    | Ratio | RatioSD | Gen0        | Gen1       | Gen2      | Allocated    | Alloc Ratio |
|-------------------- |---- |------------:|-----------:|----------:|------:|--------:|------------:|-----------:|----------:|-------------:|------------:|
| ClosedXml           | 1   |    59.84 ms | 130.093 ms |  7.131 ms |  1.00 |    0.00 |   1666.6667 |          - |         - |   8076.14 KB |        1.00 |
| FakeExcelSerializer | 1   |    19.86 ms |   2.429 ms |  0.133 ms |  0.34 |    0.04 |           - |          - |         - |    125.19 KB |        0.02 |
|                     |     |             |            |           |       |         |             |            |           |              |             |
| ClosedXml           | 10  |   404.73 ms | 265.054 ms | 14.528 ms |  1.00 |    0.00 |  16000.0000 |  1000.0000 |         - |  76114.77 KB |       1.000 |
| FakeExcelSerializer | 10  |    32.11 ms | 136.936 ms |  7.506 ms |  0.08 |    0.02 |    156.2500 |          - |         - |    659.63 KB |       0.009 |
|                     |     |             |            |           |       |         |             |            |           |              |             |
| ClosedXml           | 100 | 3,821.12 ms | 362.274 ms | 19.857 ms |  1.00 |    0.00 | 149000.0000 | 11000.0000 | 2000.0000 | 748791.17 KB |       1.000 |
| FakeExcelSerializer | 100 |    75.68 ms |  12.662 ms |  0.694 ms |  0.02 |    0.00 |   1333.3333 |          - |         - |    6004.6 KB |       0.008 |

## Example-1
If you pass an object, it will be converted to an Excel file.  
![image](https://user-images.githubusercontent.com/16958552/185727609-79b574e8-b40c-46dc-83c9-74b078a1f44a.png)
~~~
ExcelSerializer.ToFile(new string[] { "test", "test2" }, @"c:\test\test.xlsx", ExcelSerializerOptions.Default);
~~~

## Example-2
Passing a class expands the property into a column.  
![image](https://user-images.githubusercontent.com/16958552/185727657-3e41dea7-1af4-4a52-99bd-1457f895b564.png)
~~~
public class Portal
{
    public string Name { get; set; }
    public string Owner { get; set; }
    public int Level { get; set; }
}

var potals = new Portal[] {
    new Portal { Name = "Portal1", Owner = "panda728", Level = 8 },
    new Portal { Name = "Portal2", Owner = "panda728", Level = 1 },
    new Portal { Name = "Portal3", Owner = "panda728", Level = 2 },
};

ExcelSerializer.ToFile(potals, @"c:\test\potals.xlsx", ExcelSerializerOptions.Default);
~~~
## Example-3
By setting attributes on the class, you can specify the name of the title or change the order of the columns.  
![image](https://user-images.githubusercontent.com/16958552/187447183-1c0af135-8407-4c79-be8d-0b4875973a79.png)
~~~
public class Portal
{
    [DataMember(Name = "Name Ex", Order = 3)]
    public string Name { get; set; }
    [DataMember(Name = "Owner Ex", Order = 1)]
    public string Owner { get; set; }
    [DataMember(Name = "Level Ex", Order = 2)]
    public int Level { get; set; }
}

var potals = new Portal[] {
    new Portal { Name = "Portal1", Owner = "panda728", Level = 8 },
    new Portal { Name = "Portal2", Owner = "panda728", Level = 1 },
    new Portal { Name = "Portal3", Owner = "panda728", Level = 2 },
};

var newConfig = ExcelSerializerOptions.Default with
{
    HasHeaderRecord = true,
};
ExcelSerializer.ToFile(potals, @"c:\test\potalsEx.xlsx", newConfig);
~~~
## Example-4
Options can be set to display a title line and automatically adjust column widths.  
![image](https://user-images.githubusercontent.com/16958552/185727708-18201283-bb0b-46ba-a413-dbe34c20f3a3.png)
~~~
var newConfig = ExcelSerializerOptions.Default with
{
    CultureInfo = CultureInfo.InvariantCulture,
    HasHeaderRecord = true,
    HeaderTitles = new string[] { "Name", "Owner", "Level" },
    AutoFitColumns = true,
};
ExcelSerializer.ToFile(potals, @"c:\test\potalsOp.xlsx", newConfig);
~~~

## Example-5
Optionally supports Autofilter.  
~~~
var newConfig = ExcelSerializerOptions.Default with
{
    CultureInfo = CultureInfo.InvariantCulture,
    HasHeaderRecord = true,
    HeaderTitles = new string[] { "Name", "Owner", "Level" },
    AutoFitlter = true,
};
ExcelSerializer.ToFile(potals, @"c:\test\potalsOp.xlsx", newConfig);
~~~

## Note
For the method of retrieving values from IEnumerable\<T\>, Cysharp's WebSerializer method is used.

　https://github.com/Cysharp/WebSerializer
  
The following page provides information on how to return to OpenOfficeXml.

　https://gist.github.com/iso2022jp/721df3095f4df512bfe2327503ea1119

　https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e
 
## License
This library is licensed under the MIT License.
