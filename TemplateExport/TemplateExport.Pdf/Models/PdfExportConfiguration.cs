using PdfSharp;

namespace TemplateExport.Pdf.Models
{
    public class PdfExportConfiguration
    {
        private PdfExportConfiguration()
        {
        }

        internal string TemplatePath { get; private set; }

        internal string[] TemplateHead { get; private set; } = [];

        internal string[] TemplateBody { get; private set; } = [];

        internal Dictionary<string, object> DataSets { get; } = new();

        internal string OutputPath { get; private set; }

        internal Stream OutputStream { get; private set; }

        internal PageSize PageSize { get; private set; } = PageSize.A4;

        internal string TemplateStringStartsWith { get; private set; } = "{{";

        internal string TemplateStringEndsWith { get; private set; } = "}}";

        internal string TemplateStringSeparator { get; private set; } = "::";

        internal string ForAttribute { get; private set; } = "template-for";

        internal string IfAttribute { get; private set; } = "template-if";

        internal bool PreserveRowHeight { get; private set; } = true; // Can be slow for a large number of rows

        internal bool PreserveColumnWidth { get; private set; } = true; // Can be slow for a large number of columns

        internal bool PreserveMergeCells { get; private set; } = true; // Can be slow for a large number of merged cells

        internal bool PreserveCellStyles { get; private set; } = true;

        internal bool AutoFitColumns { get; private set; }

        internal bool AutoFitRows { get; private set; }

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
            /// Sets the page size for the output.
            /// </summary>
            /// <param name="pageSize"></param>
            /// <returns></returns>
            public Builder UsePageSize(PageSize pageSize)
            {
                _config.PageSize = pageSize;
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