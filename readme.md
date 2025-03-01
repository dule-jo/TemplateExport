# TemplateExport

TemplateExport is a C# library for exporting HTML templates to PDF and Excel templates to XLSX

## Features

- Support for custom fonts and styles
- Easy integration with existing projects

# Excel

### Installation

To install TemplateExport Pdf, use the NuGet Package Manager Console:

```shell
Install-Package TemplateExport.Pdf

### Usage

using System.Collections.Generic;
using TemplateExport.Pdf.Internals;
using TemplateExport.Pdf.Models;

var 

var config = PdfExportConfiguration.CreateBuilder()
    .UseTemplatePath("path/to/template.html")
    .UseOutputPath("path/to/output.pdf")
    .AddDataSet("Person", person)
    .Build();

var exporter = new TemplateExportPdf();
exporter.Export(config);


