using System.Collections;
using System.Reflection;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTemplateExport.Utilities;

namespace ExcelTemplateExport;

public class ExcelTemplateExport : IExcelTemplateExport
{
    private ExportConfiguration _config;
    
    public void Export(ExportConfiguration config)
    {
        _config = config;
        
        var workbook = new XLWorkbook(_config.TemplatePath);
        if (workbook == null) throw new FileNotFoundException();

        foreach (var sheet in workbook.Worksheets)
        {
            var rowsWithList = new List<Tuple<IXLRow, IEnumerable<object>>>();
            foreach (var row in sheet.Rows())
            {
                IEnumerable<object> listProperty = null;
                foreach (var cell in row.Cells())
                {
                    var fieldInfo = new FieldInfo(cell.GetValue<string>());
                    if (fieldInfo.ObjectName == null) continue;

                    if (!TryGetType(fieldInfo, out var obj, out var type)) continue;

                    if (typeof(IEnumerable).IsAssignableFrom(type) && type != typeof(string))
                    {
                        if (fieldInfo.Aggregation != null)
                        {
                            var aggregateValue = (obj as IEnumerable<object>).GetAggregationValue(fieldInfo);
                            CellUtility.ChangeFromObject(type, fieldInfo, aggregateValue, cell);
                        }
                        else
                        {
                            if (listProperty != null && listProperty != obj) throw new Exception("Only one list type is allowed per row.");
                            listProperty = obj as IEnumerable<object>;
                        }
                    }
                    else CellUtility.ChangeFromObject(type, fieldInfo, obj, cell);
                }

                if (listProperty != null) rowsWithList.Add(new Tuple<IXLRow, IEnumerable<object>>(row, listProperty));
            }

            foreach (var row in rowsWithList)
            {
                ChangeFromList(row.Item1, row.Item2);
            }
        }

        if (_config.OutputStream != null) workbook.SaveAs(_config.OutputStream);
        else workbook.SaveAs(_config.OutputPath);
    }

    private bool TryGetType(FieldInfo fieldInfo, out object? obj, out Type? type)
    {
        if (!_config.FieldValues.TryGetValue(fieldInfo.ObjectName, out obj))
        {
            type = null;
            return false;
        }
        type = obj?.GetType() ?? typeof(string);
        return true;
    }

    private static void ChangeFromList(IXLRow row, IEnumerable<object> list)
    {
        if (list == null || !list.Any())
        {
            row.Delete();
        }

        row.InsertRowsBelow(list.Count() - 1);

        var firstObject = list.ElementAt(0);
        var type = firstObject.GetType();
        for (var i = 1; i <= list.Count() - 1; i++)
        {
            var objInList = list.ElementAt(i);
            var newRow = row.RowBelow(i);
            CellUtility.InsertValuesInRow(newRow, type, objInList, row);
        }

        CellUtility.InsertValuesInRow(row, type, firstObject);
    }
}