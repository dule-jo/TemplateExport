using System.Diagnostics;
using ExcelTemplateExport.Internals;
using ExcelTemplateExport.Models;
using ExcelTemplateExport.Test;
using TemplateExport.Pdf.Internals;
using TemplateExport.Pdf.Models;

var person = new Person { Name = "John", Age = 30, Amount = 1000 };
var persons = new List<Person>();

for (var i= 0; i< 10; i++)
{
    persons.Add(new Person { Name = $"Person {i}", Age = 30 + i, Amount = 1000 + i });
}

if (false)
    ExportExcel(person, persons);
else 
    ExportPdf(person, persons);

void ExportPdf(Person person1, List<Person> list)
{
    var config = PdfExportConfiguration.CreateBuilder()
        .UseTemplatePath("./Resources/test.html")
        .UseOutputPath("./Resources/output.html")
        .AddDataSet("Person", person1)
        .AddDataSet("Persons", list)
        .Build();
    
    var stopwatch = new Stopwatch();
    stopwatch.Start();  
    var export = new TemplateExportPdf();
    export.Export(config);
    stopwatch.Stop();
    Console.WriteLine($"Elapsed time: {stopwatch.ElapsedMilliseconds} ms");
}

void ExportExcel(Person person1, List<Person> list)
{
    var config = ExcelExportConfiguration.CreateBuilder()
        .UseTemplatePath("./Resources/test.xlsx")
        .UseOutputPath("./Resources/output.xlsx")
        .EnablePreserveMergeCells(false)
        .EnablePreserveRowHeight(false)
        .EnablePreserveColumnWidth(false)
        .EnablePreserveCellStyles(true)
        .AddDataSet("Person", person1)
        .AddDataSet("Persons", list)
        .Build();

    var stopwatch = new Stopwatch();
    stopwatch.Start();  
    var export = new TemplateExportExcel();
    export.Export(config);
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
    }
}