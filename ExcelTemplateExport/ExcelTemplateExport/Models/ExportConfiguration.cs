namespace ExcelTemplateExport
{
    public class ExportConfiguration
    {
        public string TemplatePath { get; set; }

        public Dictionary<string, object> FieldValues { get; set; }

        public string OutputPath { get; set; }
    }
}