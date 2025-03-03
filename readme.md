# TemplateExport.Pdf

TemplateExport.Pdf is a library for exporting PDF documents using HTML templates. It allows you to define templates with placeholders and dynamically replace them with data.

## Features

- Load HTML templates from files
- Replace placeholders with data
- Support for conditional rendering
- Support for lists and collections
- Combine multiple HTML files into a single PDF

## Installation

To install TemplateExport.Pdf, add the following NuGet package to your project:

```shell
dotnet add package TemplateExport.Pdf
```

## Usage

### 1. Configure Services

First, configure the services in your `Main` function or startup class:

```csharp
using Microsoft.Extensions.DependencyInjection;
using TemplateExport.Pdf.Extensions;

var services = new ServiceCollection();
services.AddTemplateExportPdf();
var serviceProvider = services.BuildServiceProvider();
```

### 2. Create a PDF Export Configuration

Create a configuration for exporting a PDF document:

```csharp
using TemplateExport.Pdf.Models;

var config = PdfExportConfiguration.CreateBuilder()
    .UseTemplatePath("./Resources/template.html")
    .UseOutputPath("./Resources/output.pdf")
    .AddDataSet("Person", new Person { Name = "John", Age = 30 })
    .Build();
```

### 3. Export the PDF

Resolve the `ITemplateExportPdf` service and call the `Export` method:

```csharp
var pdfExport = serviceProvider.GetService<ITemplateExportPdf>();
pdfExport.Export(config);
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

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.

