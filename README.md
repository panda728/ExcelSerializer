# ExcelSerializer
Convert object to Excel file (.xlsx) [Open XML SpreadsheetML File Format]

## Getting Started
Supporting platform is .NET Standard 2.0, 2.1, .NET 6., .NET 8.

## Usage
You can use `ExcelSerializer.ToFile` to create .xlsx file.

~~~
ExcelSerializer.ToFile(Users, "test.xlsx", ExcelSerializerOptions.Default);
~~~

## Notice

Folder creation permissions are required since a working folder will be used.

## Benchmark
| Method             | N   | Mean        | Error      | StdDev   | Ratio | RatioSD | Gen0        | Gen1       | Gen2      | Allocated    | Alloc Ratio |
|------------------- |---- |------------:|-----------:|---------:|------:|--------:|------------:|-----------:|----------:|-------------:|------------:|
| ClosedXml          | 1   |    46.41 ms |  54.183 ms | 2.970 ms |  1.00 |    0.00 |   1750.0000 |   500.0000 |         - |   8077.43 KB |        1.00 |
| RunExcelSerializer | 1   |    21.74 ms | 101.723 ms | 5.576 ms |  0.47 |    0.12 |           - |          - |         - |    126.79 KB |        0.02 |
|                    |     |             |            |          |       |         |             |            |           |              |             |
| ClosedXml          | 10  |   376.59 ms | 179.068 ms | 9.815 ms |  1.00 |    0.00 |  16000.0000 |  1000.0000 |         - |  76120.38 KB |       1.000 |
| RunExcelSerializer | 10  |    27.04 ms |   4.965 ms | 0.272 ms |  0.07 |    0.00 |    142.8571 |          - |         - |    672.77 KB |       0.009 |
|                    |     |             |            |          |       |         |             |            |           |              |             |
| ClosedXml          | 100 | 3,529.88 ms | 125.574 ms | 6.883 ms |  1.00 |    0.00 | 149000.0000 | 11000.0000 | 2000.0000 | 748783.16 KB |        1.00 |
| RunExcelSerializer | 100 |    87.16 ms |  93.944 ms | 5.149 ms |  0.02 |    0.00 |   1500.0000 |   833.3333 |         - |   9527.74 KB |        0.01 |

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
