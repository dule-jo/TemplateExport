using System.Diagnostics;
using ExcelTemplateExport;

var person = new MyClass { Name = "John", Age = 30, Amount = 1000 };
var persons = new List<MyClass>();

for (var i= 0; i< 100000; i++)
{
    persons.Add(new MyClass { Name = $"Person {i}", Age = 30 + i, Amount = 1000 + i });
}

var config = new ExportConfiguration()
{
    TemplatePath = "/home/dulejo/Desktop/Template.xlsx",
    FieldValues = new Dictionary<string, object>
    {
        { "Person", person },
        { "Persons", persons }
    },
    OutputPath = "/home/dulejo/Desktop/Output.xlsx"
};

var stopwatch = new Stopwatch();
stopwatch.Start();
var export = new ExcelTemplate();
// export.ExportOpenXml(config);
export.Export(config);
stopwatch.Stop();
Console.WriteLine($"Elapsed time: {stopwatch.ElapsedMilliseconds} ms");

public class MyClass
{
    public string Name { get; set; }

    public int Age { get; set; }

    public double Amount { get; set; }
}