# EPPlus.TableAsEnumerable
Generic extension method enabling retrival of ExcelTable rows as an enumeration of typed objects.

## Synopsis

This project adds an extension method to EPPlus ExcelTable objects that enables generic retrival of the data within. Solution supports numeric types, strings, nullables and enums. It is not enabling nor supporting changing of the data in the Excel table. 

## Complex code example 

*Note:* This example is used in test case `ComplexExampleTest`, that you can find in the `EPPlus.TableAsEnumerable.Tests` project. 

Let's suppose we have the following table in a worksheet:

| License plate | Manufacturer | Manufacturing date | Price | Color | Is ready for traffic? |
|---------------|--------------|--------------------|-------|--------|-----------------------|
| A-000123 | Ford | 2001.12.20 | 3500 | Red | IGAZ |
| ACX-224 | FORD | 2010.07.30 | 4800 | Blue | IGAZ |
| EU-225541 | Opel | 2015.03.10 | 12000 | Green | IGAZ |
|  | Toyota | 1990.06.06 | 500 | Gray | HAMIS |
| ABD-002 | 2 |  | 1200 | Orance | IGAZ |

(Please note the missing values and that I have used some locales - as when recognized Excel will store them locale agnostic and IGAZ will be stored as *true*, while HAMIS as *false*).

And as developer of a more complex application we need to read this data.

Let's suppose we would like to have it in this class:

```cs
enum Manufacturers { Opel = 1, Ford, Toyota };
	
class Cars
{
	public string licensePlate { get; set; }
	public Manufacturers manufacturer { get; set; }
	public DateTime? manufacturingDate { get; set; }
	public int price { get; set; }
	public Color color { get; set; }
	public bool ready { get; set; }
	public override string ToString()
	{
		return $"{(color.ToString())} {(manufacturer.ToString())} {(manufacturingDate?.ToShortDateString())}";
	}
	
	/* 
	* other methods and properties 
	*/
}
```

With this library it is as simple as decorating the properties with `ExcelTableColumn` attribute. Like this:
```cs
{
	[ExcelTableColumn(ColumnIndex = 1)]
	public string licensePlate { get; set; }

	[ExcelTableColumn]
	public Manufacturers manufacturer { get; set; }

	[ExcelTableColumn(ColumnName = "Manufacturing date")]
	public DateTime? manufacturingDate { get; set; }

	[ExcelTableColumn]
	public int price { get; set; }

	[ExcelTableColumn]
	public Color color { get; set; }

	[ExcelTableColumn(ColumnName = "Is ready for traffic?")]
	public bool ready { get; set; }
    
    /*
    * Rest of the class
    */
}
```

And than, when we have our table in EPPLus as ExcelTable, we can just enumerate the rows as objects:
```cs
var table = excelPackage.Workbook.Worksheets["Cars"].Tables["Cars"];

foreach(var car in table.AsEnumerable<Cars>())
{
	Console.WriteLine(car);
}
```

## API reference
Code is troughout documented, but here are the key elements one needs to know when using this library.

### ExcelTableColumnAttribute
As can be noticed above, this attribute denotes property-column mapping. `ColumnName` parameter canbe used to map by column name, while `ColumnIndex` uses the *n*th one-based column. If no parameter is added, the mapping is done using the name of the decorated property.
Both parameters can not be given. Empty column name is also not accepted.

### ExcelTableConvertException and ExcelTableConvertExceptionArgs
This cutom exception is thrown when setting property to a property is failing and the extension method is told not to skip errors (this is the default case). The exception object has an args property of `ExcelTableConvertExceptionArgs` that will hold the exact circumstances of the conversion error, including the original exception as inner exception.

### AsEnumerable extension method
This generic method is doing the job, as can be seen in the example above. It returns an IEnumerable, which means that it is executed only when enumerated. Thus you might get well trough ihis call and get the exception when iterating or converting the result. 

**Note:** only classes with parameterless constructor can be used as generating type.

### Version history
**v1.0**
* Every simple type can be mapped including numeric ones, bool, string, DateTime and enumerations.
* Enumerations with underlaying type other than int are supported.
* Both value and string representation of enumeration elements can be retrieved (even mixed)
* Enums and column names are case insensitive
* Nullable properties are supported
* Complex type mappings are not suported *(might not be supported in the future either)*
* It is not necessary to map all properties, and it is not necessary to map properties to all columns either. One column can be mapped to more than one property, even uder different type.
* Mapping can be done by either the name or the index of the column. Automatic mapping to the property name is also possible.
* .NET 4.6 is needed because of some reflection properties  


## Motivation

[EPPlus](http://epplus.codeplex.com/) is a great tool allowing manipulation of Excel files of *xlsx* format files without the need of installed Excel instance and interop, by directly manipulating the OpenXML format. This is actually the only viable approach in many cases. 

EPPlus lacks of some features, for example strongly typed access of data. This project aims to fill this gap partially by providing a way to read table objects stored in the worksheets as an enumerable of typed objects. 

## Installation

* Clone it, build it, use it!
* Or get it from NuGet (https://www.nuget.org/packages/zorgoz.EPPlus.TableAsEnumerable/)
* Or to install EPPlus.TableAsEnumerable, run the following command in the Package Manager Console: `Install-Package zorgoz.EPPlus.TableAsEnumerable`

## Tests

See included `EPPlus.TableAsEnumerable.Tests` test project.

## Contributors

Zoltán Zörgő (for now, contributors are welcome)

## License

Project is licensed under MIT terms
