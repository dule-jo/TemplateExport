using System.Diagnostics;
using ExcelTemplateExport.Internals;
using ExcelTemplateExport.Models;
using ExcelTemplateExport.Test;
using TemplateExport.Pdf.Internals;
using TemplateExport.Pdf.Models;

var person = new Person { Name = "John", Age = 29, Amount = 1000 };
var persons = new List<Person>();

for (var i = 0; i < 3; i++)
{
    persons.Add(new Person { Name = $"Person {i}", Age = 30 + i, Amount = 1000 + i });
}

if (false)
    ExportExcel(person, persons);
else if (false)
    ExportPdf(person, persons);
else ExportPdf2(person, persons);

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

void ExportPdf2(Person person1, List<Person> list)
{
    var config = PdfExportConfiguration.CreateBuilder()
        .UseTemplateHead("./Resources/head1.html")
        .UseTemplateHead("./Resources/head2.html")
        .UseTemplateBody("./Resources/body1.html")
        .UseTemplateBody("./Resources/body2.html")
        .UseOutputPath("./Resources/headbody.pdf")
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

        public bool IsOld => Age > 31;
    }
}