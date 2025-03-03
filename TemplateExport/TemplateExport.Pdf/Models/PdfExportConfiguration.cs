using iText.Kernel.Geom;

namespace TemplateExport.Pdf.Models
{
    public class PdfExportConfiguration
    {
        private PdfExportConfiguration() { }
        
        internal string TemplatePath { get; set; }

        internal string[] TemplateHead { get; set; } = [];
        
        internal string[] TemplateBody { get; set; } = [];

        internal Dictionary<string, object> DataSets { get; } = new();

        internal string OutputPath { get; set; }

        internal Stream OutputStream { get; set; }

        internal string TemplateStringStartsWith { get; set; } = "{{";
        
        internal string TemplateStringEndsWith { get; set; } = "}}";
        
        internal string TemplateStringSeparator { get; set; } = "::";
        
        internal string ForAttribute { get; set; } = "template-for";
        
        internal string IfAttribute { get; set; } = "template-if";
        
        internal string ElseAttribute { get; set; } = "template-else";

        internal bool PreserveRowHeight { get; set; } = true; // Can be slow for a large number of rows
        
        internal bool PreserveColumnWidth { get; set; } = true; // Can be slow for a large number of columns

        internal bool PreserveMergeCells { get; set; } = true; // Can be slow for a large number of merged cells
        
        internal bool PreserveCellStyles { get; set; } = true;
        
        internal bool AutoFitColumns { get; set; }
        
        internal bool AutoFitRows { get; set; }
        
        internal PageSize PageSize { get; set; } = PageSize.A4;
        
        internal PageOrientation PageOrientation { get; set; } = PageOrientation.Portrait;
        
        public static Builder CreateBuilder() => new();

        public class Builder
        {
            private readonly PdfExportConfiguration _config = new PdfExportConfiguration();
            
            public PdfExportConfiguration Build() => _config;

            /// <summary>
            /// Sets the path to the template file.
            /// </summary>
            /// <param name="templatePath"></param>
            /// <returns></returns>
            public Builder UseTemplatePath(string templatePath)
            {
                _config.TemplatePath = templatePath;
                return this;
            }
            
            /// <summary>
            /// Adds a template head to the configuration. Heads are added to html file in the order they are called.
            /// Content of file must be valid html element
            /// </summary>
            /// <param name="templateHead">Path to template file that contains part or whole head</param>
            /// <returns></returns>
            public Builder UseTemplateHead(string templateHead)
            {
                _config.TemplateHead = _config.TemplateHead.Append(templateHead).ToArray();
                return this;
            }
            
            /// <summary>
            /// Adds a template body to the configuration. Bodies are added to html file in the order they are called.
            /// </summary>
            /// <param name="templateBody">Path to template file that contains part or whole header</param>
            /// <returns></returns>
            public Builder UseTemplateBody(string templateBody)
            {
                _config.TemplateBody = _config.TemplateBody.Append(templateBody).ToArray();
                return this;
            }
            
            /// <summary>
            /// sets the path to the output file. Only set this if output should be a file.
            /// </summary>
            /// <param name="outputPath"></param>
            /// <returns></returns>
            public Builder UseOutputPath(string outputPath)
            {
                _config.OutputPath = outputPath;
                return this;
            }
            
            /// <summary>
            /// Sets the output stream. Only set this if output should be a stream.
            /// </summary>
            /// <param name="outputStream"></param>
            /// <returns></returns>
            public Builder UseOutputStream(Stream outputStream)
            {
                _config.OutputStream = outputStream;
                return this;
            }
            
            /// <summary>
            /// Sets the start of the template string.
            /// </summary>
            /// <param name="templateStringStartsWith">Default is "{{"</param>
            /// <returns></returns>
            public Builder UseTemplateStringStartsWith(string templateStringStartsWith)
            {
                _config.TemplateStringStartsWith = templateStringStartsWith;
                return this;
            }
            
            /// <summary>
            /// sets the end of the template string.
            /// </summary>
            /// <param name="templateStringEndsWith">Default is "}}"</param>
            /// <returns></returns>
            public Builder UseTemplateStringEndsWith(string templateStringEndsWith)
            {
                _config.TemplateStringEndsWith = templateStringEndsWith;
                return this;
            }
            
            /// <summary>
            /// Sets the separator for the template string.
            /// </summary>
            /// <param name="templateStringSeparator">Default is "::"</param>
            /// <returns></returns>
            public Builder UseTemplateStringSeparator(string templateStringSeparator)
            {
                _config.TemplateStringSeparator = templateStringSeparator;
                return this;
            }
            
            /// <summary>
            /// Add a data set to the configuration.
            /// </summary>
            /// <param name="key"></param>
            /// <param name="value"></param>
            /// <returns></returns>
            public Builder AddDataSet(string key, object value)
            {
                _config.DataSets.Add(key, value);
                return this;
            }
            
            /// <summary>
            /// Sets the page size for the output.
            /// </summary>
            /// <param name="pageSize">Default is A4</param>
            /// <returns></returns>
            public Builder SetPageSize(PageSize pageSize)
            {
                _config.PageSize = pageSize;
                return this;
            }
            
            /// <summary>
            /// Sets the page orientation for the output.
            /// </summary>
            /// <param name="pageOrientation">Default is Portrait</param>
            /// <returns></returns>
            public Builder SetPageOrientation(PageOrientation pageOrientation)
            {
                _config.PageOrientation = pageOrientation;
                return this;
            }
            
            /// <summary>
            /// Add multiple data sets to the configuration.
            /// </summary>
            /// <param name="dataSets"></param>
            /// <returns></returns>
            public Builder AddDataSets(Dictionary<string, object> dataSets)
            {
                foreach (var dataSet in dataSets)
                {
                    _config.DataSets.Add(dataSet.Key, dataSet.Value);
                }
                return this;
            }
        }
    }
}