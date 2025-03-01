using System.Diagnostics;
using ExcelTemplateExport.Internals;
using ExcelTemplateExport.Models;
using ExcelTemplateExport.Test;
using TemplateExport.Pdf.Internals;
using TemplateExport.Pdf.Models;

var person = new Person { Name = "John", Age = 122, Amount = 1000 };
var people = new List<Person>();

for (var i = 0; i < 3; i++)
{
    people.Add(new Person { Name = $"Person {i}", Age = 30 + i, Amount = 1000 + i });
}

if (false)
    ExportExcel(person, people);
else if (false)
    ExportPdf(person, people);
else ExportPdf2(person, people);

void ExportPdf(Person person1, List<Person> list)
{
    var config = PdfExportConfiguration.CreateBuilder()
        .UseTemplatePath("./Resources/test.html")
        .UseOutputPath("./Resources/output.html")
        .AddDataSet("Person", person1)
        .AddDataSet("People", list)
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
        .AddDataSet("People", list)
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
        .AddDataSet("People", list)
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