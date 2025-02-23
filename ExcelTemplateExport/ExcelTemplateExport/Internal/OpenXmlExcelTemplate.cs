using System.Collections;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTemplateExport.Utilities;

namespace ExcelTemplateExport
{
    public class OpenXmlExcelTemplate : IExcelTemplateExport
    {
        public void Export(ExportConfiguration config)
        {
            using var template = SpreadsheetDocument.Open(config.TemplatePath, true);
            using var document = SpreadsheetDocument.Create(config.OutputPath, template.DocumentType);

            var workbookPart = document.WorkbookPart;
            var workbook = workbookPart.Workbook;

            var sheets = workbook.Descendants<Sheet>();
            foreach (var sheet in sheets) WriteDataInSheet(workbookPart, sheet, config);

            workbookPart.Workbook.Save();
            document.Save();
        }

        private static void WriteDataInSheet(WorkbookPart workbookPart, Sheet sheet, ExportConfiguration config)
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            var sharedStringPart = workbookPart.SharedStringTablePart;
            var values = sharedStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();

            var rows = worksheetPart.Worksheet.Descendants<Row>();
            var listProperty = default(IEnumerable<object>);
            var rowsWithList = new List<Tuple<Row, IEnumerable<object>>>();
            foreach (var row in rows)
            {
                var cells = row.Descendants<Cell>();
                var valuesInRow = cells.Select(x => CellUtility.GetCellValue(values, x));
                if (!valuesInRow.Any(CellUtility.IsMeantForReplace)) continue;
                foreach (var cell in cells)
                {
                    var oldValue = CellUtility.GetCellValue(values, cell);
                    if (!CellUtility.IsMeantForReplace(oldValue)) continue;

                    var fieldInfo = new FieldInfo(oldValue);
                    if (fieldInfo == null) continue;

                    if (!config.FieldValues.TryGetValue(fieldInfo.ObjectName, out var obj)) continue;
                    var type = obj.GetType();

                    if (typeof(IEnumerable).IsAssignableFrom(type))
                    {
                        if (listProperty != null && listProperty != obj) throw new Exception("Only one list type is allowed per row.");
                        listProperty = obj as IEnumerable<object>;
                    }
                    else ChangeFromObject(type, fieldInfo.PropertyName, obj, cell);
                }

                if (listProperty == null) continue;
                rowsWithList = rowsWithList.Prepend(new Tuple<Row, IEnumerable<object>>(row, listProperty)).ToList(); // Prepend so that rows are processed from the bottom up to avoid index issues
            }

            foreach (var change in rowsWithList)
            {
                var row = change.Item1;
                var list = change.Item2;
                var sheetData = row.Parent as SheetData;

                foreach (var obj in list.Take(1))
                {
                    var newRow = (Row)row.CloneNode(true);

                    var cells = newRow.Descendants<Cell>();
                    foreach (var cell in cells)
                    {
                        var value = CellUtility.GetCellValue(values, cell);
                        if (!CellUtility.IsMeantForReplace(value)) continue;

                        var fieldInfo = new FieldInfo(value);
                        if (fieldInfo == null) continue;

                        var property = fieldInfo.PropertyName;
                        ChangeFromObject(obj.GetType(), property, obj, cell);
                    }
                    row.InsertAfterSelf(newRow);
                }
            }
        }

        private static void ChangeFromObject(Type type, string propertyName, object value, Cell cell)
        {
            var property = type.GetProperty(propertyName);
            if (property == null) return;
            var newValue = type.GetProperty(propertyName)?.GetValue(value);

            cell.InsertValue(newValue);
        }
    }
}