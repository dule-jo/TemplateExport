using System.Diagnostics;
using ExcelTemplateExport;
using ExcelTemplateExport.Extensions;
using ExcelTemplateExport.Models;
using ExcelTemplateExport.Test;
using iText.Kernel.Geom;
using Microsoft.Extensions.DependencyInjection;
using TemplateExport.Pdf;
using TemplateExport.Pdf.Extensions;
using TemplateExport.Pdf.Models;

var services = new ServiceCollection();
services.AddTemplateExportExcel();
services.AddTemplateExportPdf();
var serviceProvider = services.BuildServiceProvider();
var excelExport = serviceProvider.GetService<ITemplateExportExcel>();
var pdfExport = serviceProvider.GetService<ITemplateExportPdf>();

var person = new Person { Name = "John", Age = 122, Amount = 1000 };
var people = new List<Person>();

for (var i = 0; i < 3; i++)
{
    people.Add(new Person { Name = $"Person {i}", Age = 30 + i, Amount = 1000 + i });
}

ExportExcel(person, people, excelExport);
// ExportPdf(person, people, pdfExport);
// ExportPdf2(person, people, pdfExport);

void ExportPdf(Person person1, List<Person> list, ITemplateExportPdf templateExportPdf)
{
    var stopwatch = new Stopwatch();
    stopwatch.Start();
    
    var config = PdfExportConfiguration.CreateBuilder()
        .UseTemplatePath("./Resources/test.html")
        .UseOutputPath("./Resources/output.pdf")
        .AddDataSet("Person", person1)
        .AddDataSet("People", list)
        .Build();
    
    templateExportPdf.Export(config);
    stopwatch.Stop();
    
    Console.WriteLine($"Elapsed time: {stopwatch.ElapsedMilliseconds} ms");
}

void ExportPdf2(Person person1, List<Person> list, ITemplateExportPdf templateExportPdf)
{

    var stopwatch = new Stopwatch();
    stopwatch.Start();
    
    var config = PdfExportConfiguration.CreateBuilder()
        .UseTemplateHead("./Resources/head1.html")
        .UseTemplateHead("./Resources/head2.html")
        .UseTemplateBody("./Resources/body1.html")
        .UseTemplateBody("./Resources/body2.html")
        .UseOutputPath("./Resources/headbody.pdf")
        .SetPageOrientation(PageOrientation.Landscape)
        .SetPageSize(PageSize.A3)
        .AddDataSet("Person", person1)
        .AddDataSet("People", list)
        .Build();
    
    templateExportPdf.Export(config);
    stopwatch.Stop();
    
    Console.WriteLine($"Elapsed time: {stopwatch.ElapsedMilliseconds} ms");
}

void ExportExcel(Person person, List<Person> list, ITemplateExportExcel templateExportExcel)
{
    var stopwatch = new Stopwatch();
    stopwatch.Start();
    
    var config = ExcelExportConfiguration.CreateBuilder()
        .UseTemplatePath("./Resources/test.xlsx")
        .UseOutputPath("./Resources/output.xlsx")
        .EnablePreserveMergeCells(true)
        // .EnablePreserveRowHeight(true)
        // .EnablePreserveColumnWidth(true)
        .EnableAutoFitColumns(true)
        .EnablePreserveCellStyles(true)
        .AddDataSet("Person", person)
        .AddDataSet("Persons", list)
        .Build();
    
    templateExportExcel.Export(config);
    stopwatch.Stop();
    
    Console.WriteLine($"Elapsed time: {stopwatch.ElapsedMilliseconds} ms");
}

namespace ExcelTemplateExport.Test
{
    public class Person
    {
        public string Name { get; set; }

        public int Age { get; set; }

        public double Amount { get; set; }

        public double AmountPerYear => Amount / Age;

        public bool IsOld => Age > 31;
    }
}