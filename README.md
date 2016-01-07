# LFEG - Lighitng Fast Excel Generator

The library was built to generate big excel reports which don't need fancy formatting or any formulas. It can be compared to export to CSV but if you need that data to open in Excel and do something with that data (sort, aggregate), it will be very helpful to have correct column data types but CSV doesn't have types, everythin there is actually a string.

There are .net libraries which allow to create Excel files but the ones I saw were very-very slow when you have a lot of data. But this library allows to generate files with 5000 rows in 200ms or so depending on hardware. And even reports with 1 000 000 rows of data are not a problem, it will take less than half of a minute to generate, or more depending on your hardware and data.

## Usage
Add attributes to your data models
```
class Model
{
    [IgnoreExcelExport]
    public Guid Id { get; set; }

    [ExcelExport(Caption = "Title hello", Width = 40)]
    public string TitleInline { get; set; }

    [ExcelExport(Intern = true)]
    public string TitleShared { get; set; }

    [ExcelExport(DataFormat = "\"Yes\";;\"No\"")]
    public bool IsEnabled { get; set; }
    
    public bool IsEnabled2 { get; set; }
    public int Width { get; set; }
    public ModelType Type { get; set; }
    public DateTime UpdatedDate { get; set; }
}
```
Setup generator
```
var builder = ExcelFileGeneratorSettings.Create()
                .SetDefaultDateFormat("dd/mm/yy")
                .SetCompressionLevel(CompressionLevel.BestCompression);
```
Generate!
```
var generator = builder.CreateGenerator();
using (var stream = File.Create("c:\\output.xlsx"))
{
	generator.GenerateFile(data, stream);
}
```

