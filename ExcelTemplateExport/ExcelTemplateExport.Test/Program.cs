using System.Diagnostics;
using ExcelTemplateExport;

var person = new Person { Name = "John", Age = 30, Amount = 1000 };
var persons = new List<Person>();

for (var i= 0; i< 100000; i++)
{
    persons.Add(new Person { Name = $"Person {i}", Age = 30 + i, Amount = 1000 + i });
}

var config = new ExportConfiguration()
{
    TemplatePath = "/home/dulejo/Desktop/aaa.xlsx",
    FieldValues = new Dictionary<string, object>
    {
        { "Person", person },
        { "Persons", persons },
        { "HeaderName", "Neki Header" }
    },
    OutputPath = "/home/dulejo/Desktop/Output.xlsx"
};

var stopwatch = new Stopwatch();
stopwatch.Start();
var export = new NewClosedXmlExcelTemplate();
export.Export(config);
stopwatch.Stop();
Console.WriteLine($"Elapsed time: {stopwatch.ElapsedMilliseconds} ms");

internal class Person
{
    public string Name { get; set; }

    public int Age { get; set; }

    public double Amount { get; set; }
}