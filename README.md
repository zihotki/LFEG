# LFEG - Lighitng Fast Excel Generator

The library was built to generate big excel reports which don't need fancy formatting or any formulas. It can be compared to export to CSV but if you need that data to open in Excel and do something with that data (sort, aggregate), it will be very helpful to have correct column data types but CSV doesn't have types, everythin there is actually a string.

There are .net libraries which allow to create Excel files but the ones I saw were very-very slow when you have a lot of data. But this library allows to generate files with 5000 rows in 200ms or so depending on hardware. And even reports with 1 000 000 rows of data are not a problem, it will take less than half of a minute to generate, or more depending on your hardware and data.

## Installation
Simply run in Package Manager Console
```
PM> Install-Package LFEG
```

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

## Data Convertion 

Excel has only the following cell types so convertion should be applied to data:

* Shared String
* Inline String
* Boolean
* Number
* Date
* Formula
* Error

Currently LFEG doesn't use Formula and Error cell types. 

Data convertion is performed by DataProviderVisitors. A property or a data column is checked by visitors one by one in the following order and visitors check data type of data and determine if they can convert the data:

* EnumDataProviderVisitor - handles all enum properties and converts enum to string using value provided in DescriptionAttribute or string value of an enum splitting the name by camel case - "FooBar" becomes "Foo Bar", splitting will be performed only for a-zA-Z characters. Shared String cell type is used for enum data (data is interned).
* NumericDataProviderVisitor - handles Byte, Decimal, Double, Int16, Int32, Int64, SByte, Single, UInt16, UInt32, UInt64 types and uses Number cell type for number data. To convert number to xml string Object.ToString method is used (this will change in the future since not all cultures use dot ('.') as decimal separator but the result value should use it). Interning can not be applied for this column.
* DateTimeDataProviderVisitor - handles DateTime type and uses Date cell type for date data. The date is converted to OLE Automation date. Interning can not be applied for this column.
* BooleanDataProviderVisitor - handles Boolean type and uses either Boolean cell type if no data format is applied for the property or as a default boolean data format, or Number otherwise, where 0 - is False, 1 - is True. Since OOXML doesn't support data formats for boolean type but sometimes users want to show Yes/No instead of TRUE/FALSE this workaround was implemented. Interning can not be applied for this column.
* DefaultStringDataProviderVisitor - handles all remaining types and uses Inline String cell type for data. Interning can be applied for this column and in that case Shared String will be used as a cell type.
