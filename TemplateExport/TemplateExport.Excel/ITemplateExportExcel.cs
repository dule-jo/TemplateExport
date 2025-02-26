using TemplateExport.Excel.Models;

namespace TemplateExport.Excel
{
    public interface ITemplateExportExcel
    {
        public void Export(ExcelExportConfiguration config);
    }
}