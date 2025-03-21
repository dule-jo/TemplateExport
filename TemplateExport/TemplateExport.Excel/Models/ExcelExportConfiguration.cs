namespace ExcelTemplateExport.Models
{
    public class ExcelExportConfiguration
    {
        private ExcelExportConfiguration() { }
        
        internal string TemplatePath { get; set; }

        internal Dictionary<string, object> DataSets { get; } = new();

        internal string OutputPath { get; set; }

        internal Stream OutputStream { get; set; }

        internal string TemplateStringStartsWith { get; set; } = "{{";
        
        internal string TemplateStringEndsWith { get; set; } = "}}";
        
        internal string TemplateStringSeparator { get; set; } = "::";

        internal bool PreserveRowHeight { get; set; } = true; // Can be slow for a large number of rows
        
        internal bool PreserveColumnWidth { get; set; } = true; // Can be slow for a large number of columns

        internal bool PreserveMergeCells { get; set; } = true; // Can be slow for a large number of merged cells
        
        internal bool PreserveCellStyles { get; set; } = true;
        
        internal bool AutoFitColumns { get; set; }
        
        internal bool AutoFitRows { get; set; }
        
        internal bool CopyGraphs { get; set; }
        
        public static Builder CreateBuilder() => new();

        public class Builder
        {
            private readonly ExcelExportConfiguration _config = new ExcelExportConfiguration();
            
            public ExcelExportConfiguration Build() => _config;

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
            /// Sets whether to preserve row height. Set false if there are a large number of rows.
            /// Automatically sets AutoFitRows to false if true.
            /// </summary>
            /// <param name="preserveRowHeight"> Default is true.</param>
            /// <returns></returns>
            public Builder EnablePreserveRowHeight(bool preserveRowHeight)
            {
                _config.PreserveRowHeight = preserveRowHeight;
                if (!preserveRowHeight) _config.AutoFitRows = false;
                return this;
            }
            
            /// <summary>
            /// Sets whether to preserve column width. Set false if there are a large number of columns.
            /// </summary>
            /// <param name="preserveColumnWidth">Default is true</param>
            /// <returns></returns>
            public Builder EnablePreserveColumnWidth(bool preserveColumnWidth)
            {
                _config.PreserveColumnWidth = preserveColumnWidth;
                if (!preserveColumnWidth) _config.AutoFitColumns = false;
                return this;
            }
            
            /// <summary>
            /// Sets whether to preserve merge cells.
            /// </summary>
            /// <param name="preserveMergeCells">Default is true</param>
            /// <returns></returns>
            public Builder EnablePreserveMergeCells(bool preserveMergeCells)
            {
                _config.PreserveMergeCells = preserveMergeCells;
                return this;
            }
            
            /// <summary>
            /// Sets whether to preserve cell styles.
            /// </summary>
            /// <param name="preserveCellStyles">Default is true</param>
            /// <returns></returns>
            public Builder EnablePreserveCellStyles(bool preserveCellStyles)
            {
                _config.PreserveCellStyles = preserveCellStyles;
                return this;
            }
            
            /// <summary>
            /// Sets whether to auto fit columns. Auto fitting columns can be slow for a large number of columns.
            /// Automatically sets PreserveColumnWidth to false if true.
            /// </summary>
            /// <param name="autoFitColumns">Default is false.</param>
            /// <returns></returns>
            public Builder EnableAutoFitColumns(bool autoFitColumns)
            {
                _config.AutoFitColumns = autoFitColumns;
                if (autoFitColumns) _config.PreserveColumnWidth = false;
                return this;
            }
            
            /// <summary>
            /// Sets whether to auto fit rows. Auto fitting rows can be slow for a large number of rows.
            /// Automatically sets PreserveRowHeight to false if true.
            /// </summary>
            /// <param name="autoFitRows"></param>
            /// <returns></returns>
            public Builder EnableAutoFitRows(bool autoFitRows)
            {
                _config.AutoFitRows = autoFitRows;
                if (autoFitRows) _config.PreserveRowHeight = false;
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

            /// <summary>
            /// Sets whether to copy graphs from the template to the output.
            /// </summary>
            /// <param name="copyGraphs"></param>
            /// <returns></returns>
            public Builder EnableCopyGraphs(bool copyGraphs = true)
            {
                _config.CopyGraphs = copyGraphs;
                return this;
            }
        }
    }
}