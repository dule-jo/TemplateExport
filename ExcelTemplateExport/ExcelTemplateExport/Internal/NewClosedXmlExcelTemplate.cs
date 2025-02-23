using System.Collections;
using ClosedXML.Excel;
using ExcelTemplateExport.Utilities;

namespace ExcelTemplateExport
{
    public class NewClosedXmlExcelTemplate
    {
        private static int rowsInserted = 0;

        public void Export(ExportConfiguration config)
        {
            var workbook = new XLWorkbook(config.TemplatePath);
            if (workbook == null) throw new FileNotFoundException();
            var outputWorkbook = new XLWorkbook();

            foreach (var sheet in workbook.Worksheets)
            {
                var outputSheet = outputWorkbook.Worksheets.Add(sheet.Name);
                var rowsWithList = new List<Tuple<IXLRow, IEnumerable<object>>>();
                foreach (var row in sheet.Rows())
                {
                    CopyRow(config, row, outputSheet, rowsWithList);
                }
            }

            outputWorkbook.SaveAs(GetOutputPath(config.OutputPath));
        }

        private static void CopyRow(ExportConfiguration config, IXLRow row, IXLWorksheet outputSheet, List<Tuple<IXLRow, IEnumerable<object>>> rowsWithList)
        {
            var listObject = IsListRow(row.Cells(), config);
            if (listObject != null)
            {
                foreach (var member in listObject)
                {
                    CopyCells(config, row, outputSheet, listObject, member);
                    if (member != listObject.LastOrDefault())
                        rowsInserted++;
                }
            }
            else
            {
                CopyCells(config, row, outputSheet, listObject, null);
            }
        }

        private static void CopyCells(ExportConfiguration config, IXLRow row, IXLWorksheet outputSheet, IEnumerable<object> listObject, object member)
        {
            foreach (var cell in row.Cells())
            {
                var oldValue = cell.GetValue<string>();
                var fieldInfo = new FieldInfo(oldValue);
                if (fieldInfo.ObjectName == null || fieldInfo.PropertyName == null)
                {
                    CopyValueToDestination(cell, oldValue, outputSheet);
                    continue;
                }

                if (!config.FieldValues.TryGetValue(fieldInfo.ObjectName, out var obj)) continue;
                var type = obj.GetType();

                if (typeof(IEnumerable).IsAssignableFrom(type))
                {
                    if (member == null) continue;
                    var memberType = member.GetType();
                    var value = memberType.GetProperty(fieldInfo.PropertyName)?.GetValue(member);
                    CopyValueToDestination(cell, value?.ToString() ?? string.Empty, outputSheet);
                }
                else
                {
                    var value = type.GetProperty(fieldInfo.PropertyName)?.GetValue(obj);
                    CopyValueToDestination(cell, value?.ToString() ?? string.Empty, outputSheet);
                }
            }
        }

        private static void CopyValueToDestination(IXLCell cell, string value, IXLWorksheet outputSheet) { outputSheet.Cell(cell.Address.RowNumber + rowsInserted, cell.Address.ColumnNumber).InsertValue(value); }

        private static string GetOutputPath(string configOutputPath) => string.IsNullOrEmpty(configOutputPath) ? "./output.xlsx" : configOutputPath;

        private static IEnumerable<object> IsListRow(IEnumerable<IXLCell> cells, ExportConfiguration config)
        {
            IEnumerable<object> response = null;
            foreach (var cell in cells)
            {
                if (cell.DataType == XLDataType.Blank) continue;
                var oldValue = cell.GetText();
                var fieldInfo = new FieldInfo(oldValue);
                if (fieldInfo.ObjectName == null || fieldInfo.PropertyName == null) continue;
                if (!config.FieldValues.TryGetValue(fieldInfo.ObjectName, out var obj)) continue;
                var type = obj.GetType();
                if (typeof(IEnumerable).IsAssignableFrom(type))
                {
                    if (obj is IEnumerable<object> && response != null && response != obj as IEnumerable<object>)
                    {
                        throw new Exception("Only one list type is allowed per row.");
                    }

                    response = obj as IEnumerable<object>;
                }
            }
            return response;
        }
    }
}