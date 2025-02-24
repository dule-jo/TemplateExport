namespace ExcelTemplateExport
{
    public class ExportConfiguration
    {
        public string TemplatePath { get; set; }

        public Dictionary<string, object> FieldValues { get; set; }

        public string OutputPath { get; set; }

        public Stream OutputStream { get; set; }

        public string TemplateStringStartsWith { get; set; } = "{{";
        
        public string TemplateStringEndsWith { get; set; } = "}}";
        
        public string TemplateStringSeparator { get; set; } = "::";
    }
}