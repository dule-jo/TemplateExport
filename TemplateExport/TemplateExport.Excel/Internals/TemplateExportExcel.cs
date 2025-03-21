using ClosedXML.Excel;
using ClosedXML.Graphics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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

            if (config.CopyGraphs)
            {
                var stream = CopyGraphs(templateWb, outputWb);
                if (config.OutputStream != null) stream.CopyTo(config.OutputStream);
                else if (!string.IsNullOrEmpty(config.OutputPath))
                {
                    using var fileStream = new FileStream(config.OutputPath, FileMode.Create, FileAccess.Write);
                    stream.CopyTo(fileStream);
                }
            }

            if (config.OutputStream != null)
                outputWb.SaveAs(config.OutputStream);
            else if (!string.IsNullOrEmpty(config.OutputPath)) outputWb.SaveAs(config.OutputPath);
            else throw new ArgumentNullException("OutputPath or OutputStream must be set");
        }

        private MemoryStream CopyGraphs(XLWorkbook templateWb, XLWorkbook outputWb)
        {
            using var templateStream = new MemoryStream();
            var outputStream = new MemoryStream();
            
            templateWb.SaveAs(templateStream);
            outputWb.SaveAs(outputStream);

            using var templateDoc = SpreadsheetDocument.Open(templateStream, false);
            using var outputDoc = SpreadsheetDocument.Open(outputStream, true);
            
            var templateWorkbookPart = templateDoc.WorkbookPart;
            var outputWorkbookPart = outputDoc.WorkbookPart;

            foreach (var templateSheet in templateWorkbookPart.Workbook.Sheets.Elements<Sheet>())
            {
                var templateSheetPart = (WorksheetPart)templateWorkbookPart.GetPartById(templateSheet.Id);
                var outputSheetPart = GetWorksheetPartByName(outputWorkbookPart, templateSheet.Name);

                if (templateSheetPart.DrawingsPart == null) continue;
                    
                var templateDrawingPart = templateSheetPart.DrawingsPart;
                var outputDrawingPart = outputSheetPart.AddNewPart<DrawingsPart>();
                outputDrawingPart.FeedData(templateDrawingPart.GetStream());

                foreach (var templateChartPart in templateDrawingPart.ChartParts)
                {
                    var outputChartPart = outputDrawingPart.AddNewPart<ChartPart>();
                    outputChartPart.FeedData(templateChartPart.GetStream());

                    foreach (var rel in templateChartPart.Parts)
                    {
                        outputChartPart.AddPart(rel.OpenXmlPart, rel.RelationshipId);
                    }
                }

                // 🔹 Kopiranje svih relacija ChartPart-a (ExternalData, Images, itd.)
                foreach (var externalRel in templateDrawingPart.ExternalRelationships)
                {
                    outputDrawingPart.AddExternalRelationship(externalRel.RelationshipType, externalRel.Uri);
                }

                outputSheetPart.Worksheet.AppendChild(new Drawing { Id = outputSheetPart.GetIdOfPart(outputDrawingPart) });
                outputSheetPart.Worksheet.Save();
            }

            outputStream.Position = 0;
            return outputStream;
        }

        private static WorksheetPart GetWorksheetPartByName(WorkbookPart workbookPart, string sheetName)
        {
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            return sheet != null ? (WorksheetPart)workbookPart.GetPartById(sheet.Id) : null;
        }
    }
}