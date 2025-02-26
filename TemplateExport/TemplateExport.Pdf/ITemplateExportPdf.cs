using TemplateExport.Pdf.Models;

namespace TemplateExport.Pdf;

public interface ITemplateExportPdf
{
    public void Export(PdfExportConfiguration config);
}