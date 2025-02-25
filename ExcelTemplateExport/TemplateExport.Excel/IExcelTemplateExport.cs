using ExcelTemplateExport.Models;

namespace ExcelTemplateExport
{
    public interface IExcelTemplateExport
    {
        public void Export(ExportConfiguration config);
    }
}