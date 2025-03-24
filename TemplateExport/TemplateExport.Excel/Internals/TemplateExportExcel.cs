using ClosedXML.Excel;
using ClosedXML.Graphics;
using ExcelTemplateExport.Models;

namespace ExcelTemplateExport.Internals
{
    internal class TemplateExportExcel : ITemplateExportExcel
    {
        public void Export(ExcelExportConfiguration config)
        {
            LoadOptions.DefaultGraphicEngine = new DefaultGraphicEngine("DejaVu Sans");

            using var templateWb = new XLWorkbook(config.TemplatePath);
            using var outputWb = config.CopyGraphs ? GetWorkbookWithoutData(config) : new XLWorkbook();

            foreach (var worksheet in templateWb.Worksheets)
            {
                var outputSheet = outputWb.Worksheets.FirstOrDefault(x => x.Name == worksheet.Name);
                
                if (outputSheet == null) outputSheet = outputWb.Worksheets.Add(worksheet.Name);

                new TemplateExportExcelWorksheet().Export(config, worksheet, outputSheet);
            }

            if (config.OutputStream != null)
                outputWb.SaveAs(config.OutputStream);
            else if (!string.IsNullOrEmpty(config.OutputPath)) outputWb.SaveAs(config.OutputPath);
            else throw new ArgumentNullException("OutputPath or OutputStream must be set");
        }

        private static XLWorkbook GetWorkbookWithoutData(ExcelExportConfiguration config)
        {
            var wb = new XLWorkbook(config.TemplatePath);

            foreach (var worksheet in wb.Worksheets)
            {
                foreach (var cell in worksheet.RangeUsed().Cells())
                {
                    if (!cell.HasFormula) cell.Clear(XLClearOptions.Contents);
                }
            }

            return wb;
        }
    }
}