

# TemplateExport

TemplateExport is a library for exporting PDF documents using HTML templates and exporting Excel documents using Xlsx templates. It allows you to define templates with placeholders and dynamically replace them with data.

# TemplateExport.Pdf

### Features

- Load HTML templates from files
- Replace placeholders with data
- Support for conditional rendering
- Support for lists and collections
- Combine multiple HTML files into a single PDF

### Installation

To install TemplateExport.Pdf, add the following NuGet package to your project:

```shell
dotnet add package TemplateExport.Pdf
```

### Usage

### 1. Configure Services

First, configure the services in your `Main` function or startup class:

```csharp
using Microsoft.Extensions.DependencyInjection;
using TemplateExport.Pdf.Extensions;

var services = new ServiceCollection();
services.AddTemplateExportPdf();
var serviceProvider = services.BuildServiceProvider();
```

### 2. Add html template on given location

Create html template on given location

```html
<html>
	<div template-if="{{Person::IsOld}}">
		<div template-if="{{Person::IsOld}}" style="color: red">  
		  Old: {{Person::Age}}   
		</div>  
		<div template-else style="color: green">  
		  Young: {{Person::Age}}  
		</div>
	</div>
	<div template-for="{{People}}">
		<div style="color:blue">
			{{People::Name}}
		</div>
		<div style="color:gray">
			{{People::Age}}
		</div>
	</div>
</html>

```

### 3. Create a PDF Export Configuration

Create a configuration for exporting a PDF document:

```csharp
using TemplateExport.Pdf.Models;

var config = PdfExportConfiguration.CreateBuilder()
    .UseTemplatePath("./Resources/template.html")
    .UseOutputPath("./Resources/output.pdf")
    .AddDataSet("Person", new Person { Name = "John", Age = 30 })
    .Build();
```

### 4. Export the PDF

Resolve the `ITemplateExportPdf` service and call the `Export` method:

```csharp
var pdfExport = serviceProvider.GetService<ITemplateExportPdf>();
pdfExport.Export(config);
```

### 5. Configuration options

``` csharp

UseTemplateStringStartsWith("{{") // set placeholders starts with string (Default: "{{")

UseTemplateStringEndsWith("}}") // set placeholders starts with string (Default: "}}")

UseTemplateStringSeparator("::") // set placeholders separator string (Default: "::")

UseOutputStream(stream) // set output stream for pdf

UseOutputPath(path) // set output path for pdf

AddDataSet("Person", new Person { Name = "John", Age = 30 }) // add data set to configuration

SetPageOrientation(PageOrientation.Landscape) // set page orientation (Default: Portrait)

SetPageSize(PageSize.A3) // set page size (Default: A4)

```

### Example

Here is a complete example of exporting a PDF document:

```csharp
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Extensions.DependencyInjection;
using TemplateExport.Pdf;
using TemplateExport.Pdf.Extensions;
using TemplateExport.Pdf.Models;

var services = new ServiceCollection();
services.AddTemplateExportPdf();
var serviceProvider = services.BuildServiceProvider();

var pdfExport = serviceProvider.GetService<ITemplateExportPdf>();

var person = new Person { Name = "John", Age = 30 };
var people = new List<Person>
{
    new Person { Name = "Alice", Age = 25 },
    new Person { Name = "Bob", Age = 35 }
};

var config = PdfExportConfiguration.CreateBuilder()
    .UseTemplatePath("./Resources/template.html")
    .UseOutputPath("./Resources/output.pdf")
    .AddDataSet("Person", person)
    .AddDataSet("People", people)  
	.SetPageOrientation(PageOrientation.Landscape)  
	.SetPageSize(PageSize.A3)
    .Build();

var stopwatch = new Stopwatch();
stopwatch.Start();

pdfExport.Export(config);

stopwatch.Stop();
Console.WriteLine($"Elapsed time: {stopwatch.ElapsedMilliseconds} ms");

public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
}
```


# TemplateExport.Excel

TemplateExport.Excel is a library for exporting Excel documents using templates. It allows you to define templates with placeholders and dynamically replace them with data.

## Features

- Load Excel templates from files
- Replace placeholders with data
- Support for conditional rendering
- Support for lists and collections

## Installation

To install TemplateExport.Excel, add the following NuGet package to your project:

```shell
dotnet add package TemplateExport.Excel
```

## Usage

### 1. Configure Services

First, configure the services in your `Main` function or startup class:

```csharp
using Microsoft.Extensions.DependencyInjection;
using TemplateExport.Excel.Extensions;

var services = new ServiceCollection();
services.AddTemplateExportExcel();
var serviceProvider = services.BuildServiceProvider();
```

### 2. Create excel template on given location

Create template xlsx file on location provided to configuration builder
[template-example.xlsx](./TemplateExport/TemplateExport.Test/Resources/test.xlsx)

### 3. Create an Excel Export Configuration

Create a configuration for exporting an Excel document:

```csharp
using TemplateExport.Excel.Models;

var config = ExcelExportConfiguration.CreateBuilder()
    .UseTemplatePath("./Resources/template.xlsx")
    .UseOutputPath("./Resources/output.xlsx")
    .AddDataSet("Person", new Person { Name = "John", Age = 30 })
    .Build();
```

### 4. Export the Excel

Resolve the `ITemplateExportExcel` service and call the `Export` method:

```csharp
var excelExport = serviceProvider.GetService<ITemplateExportExcel>();
excelExport.Export(config);
```

### 5. Configuration options

``` csharp

UseOutputStream(stream) // set output stream for excel

UseTemplateStringStartsWith("{{") // set placeholders starts with string (Default: "{{")

UseTemplateStringEndsWith("}}") // set placeholders starts with string (Default: "}}")

UseTemplateStringSeparator("::") // set placeholders separator string (Default: "::")

EnablePreserveRowHeight(false) // preserve row height of template file. Can be very slow for a lot of rows (Default: true)

EnablePreserveColumnWidth(false) // preserve column width of template file. Can be very slow for a lot of columns (Default: true)

EnablePreserveMergeCells(false) // preserve merge cells of template file. Can be very slow for a lot of merge cells (Default: true)

EnablePreserveCellStyles(false) // preserve cell styles of template file. Can be very slow for a lot of cells (Default: true)

EnableAutoFitColumns(true) // auto fit columns after export. Cannot be true if PreserveColumnWidth is enabled (Default: false)

EnableAutoFitRows(true) // auto fit rows after export. Cannot be true if PreserveRowHeight is enabled (Default: false)

AddDataSet("Person", new Person { Name = "John", Age = 30 }) // add data set to configuration

```

### Example

Here is a complete example of exporting an Excel document:

```csharp
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Extensions.DependencyInjection;
using TemplateExport.Excel;
using TemplateExport.Excel.Extensions;
using TemplateExport.Excel.Models;

var services = new ServiceCollection();
services.AddTemplateExportExcel();
var serviceProvider = services.BuildServiceProvider();

var excelExport = serviceProvider.GetService<ITemplateExportExcel>();

var person = new Person { Name = "John", Age = 30 };
var people = new List<Person>
{
    new Person { Name = "Alice", Age = 25 },
    new Person { Name = "Bob", Age = 35 }
};

var config = ExcelExportConfiguration.CreateBuilder()
    .UseTemplatePath("./Resources/template.xlsx")
    .UseOutputPath("./Resources/output.xlsx")
    .AddDataSet("Person", person)
    .AddDataSet("People", people)
    .Build();

var stopwatch = new Stopwatch();
stopwatch.Start();

excelExport.Export(config);

stopwatch.Stop();
Console.WriteLine($"Elapsed time: {stopwatch.ElapsedMilliseconds} ms");

public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
}
```

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.
