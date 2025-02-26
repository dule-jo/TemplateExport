using TemplateExport.Pdf.Models;

namespace TemplateExport.Pdf.Utilities
{
    internal static class PatternUtility
    {
        internal static List<FieldInfo> ExtractPatterns(string input, PdfExportConfiguration configuration)
        {
            var results = new List<FieldInfo>();
            var start = 0;

            while ((start = input.IndexOf(configuration.TemplateStringStartsWith, start)) != -1)
            {
                var end = input.IndexOf(configuration.TemplateStringEndsWith, start + configuration.TemplateStringStartsWith.Length);
                if (end == -1)
                    break;

                var match = input.Substring(start, end - start + configuration.TemplateStringEndsWith.Length);
                results.Add(new FieldInfo(match, configuration));

                start = end + 2;
            }

            return results;
        }
    }
}