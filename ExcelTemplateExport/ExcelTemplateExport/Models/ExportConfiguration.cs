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

        public bool PreserveRowHeight { get; set; } = false; // Extreamly slow for a large number of rows
        
        public bool PreserveColumnWidth { get; set; } = true; // Extreamly slow for a large number of columns
    }
}