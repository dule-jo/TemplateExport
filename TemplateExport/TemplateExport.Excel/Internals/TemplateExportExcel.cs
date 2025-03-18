using ClosedXML.Excel;
using ClosedXML.Graphics;
using ExcelTemplateExport.Models;

namespace ExcelTemplateExport.Internals
{
    internal class TemplateExportExcel : ITemplateExportExcel
    {
        public void Export(ExcelExportConfiguration config)
        {
            LoadOptions.DefaultGraphicEngine = new DefaultGraphicEngine("Arial");
            
            using var templateWb = new XLWorkbook(config.TemplatePath);
            using var outputWb = new XLWorkbook();

            foreach (var worksheet in templateWb.Worksheets)
            {
                var outputSheet = outputWb.AddWorksheet(worksheet.Name);
                new TemplateExportExcelWorksheet().Export(config, worksheet, outputSheet);
            }

            if (config.OutputStream != null)
                outputWb.SaveAs(config.OutputStream);
            else if (!string.IsNullOrEmpty(config.OutputPath)) outputWb.SaveAs(config.OutputPath);
            else throw new ArgumentNullException("OutputPath or OutputStream must be set");
        }
    }
}