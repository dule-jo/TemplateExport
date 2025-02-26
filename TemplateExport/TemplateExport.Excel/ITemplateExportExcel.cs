using ExcelTemplateExport.Models;

namespace ExcelTemplateExport
{
    public interface ITemplateExportExcel
    {
        public void Export(ExcelExportConfiguration config);
    }
}