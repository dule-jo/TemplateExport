using ClosedXML.Excel;
using ExcelTemplateExport.Models;

namespace ExcelTemplateExport.Internal
{
    public class TemplateExportExcel : IExcelTemplateExport
    {
        public void Export(ExportConfiguration config)
        {
            using var templateWb = new XLWorkbook(config.TemplatePath);
            using var outputWb = new XLWorkbook();

            foreach (var worksheet in templateWb.Worksheets)
            {
                var outputSheet = outputWb.AddWorksheet(worksheet.Name);
                new TemplateExportExcelWorksheet().Export(config, worksheet, outputSheet);
            }

            outputWb.SaveAs(config.OutputPath);
        }
    }
}