using System.Collections;
using System.Reflection;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTemplateExport.Utilities;

namespace ExcelTemplateExport;

public class ExcelTemplate : IExcelTemplateExport
{
    public void Export(ExportConfiguration config)
    {
        var workbook = new XLWorkbook(config.TemplatePath);
        if (workbook == null) throw new FileNotFoundException();

        foreach (var sheet in workbook.Worksheets)
        {
            var rowsWithList = new List<Tuple<IXLRow, IEnumerable<object>>>();
            foreach (var row in sheet.Rows())
            {
                IEnumerable<object> listProperty = null;
                foreach (var cell in row.Cells())
                {
                    var oldValue = cell.GetValue<string>();
                    if (!oldValue.StartsWith("{{") && !oldValue.EndsWith("}}")) continue;

                    oldValue = oldValue.TrimStart('{').TrimEnd('}');

                    var fieldInfo = oldValue.Split("::");
                    if (fieldInfo.Length != 2) continue;

                    var objectName = fieldInfo.ElementAt(0);
                    var propertyName = fieldInfo.ElementAt(1);

                    if (!config.FieldValues.TryGetValue(objectName, out var obj)) continue;
                    var type = obj.GetType();

                    if (typeof(IEnumerable).IsAssignableFrom(type))
                    {
                        if (listProperty != null && listProperty != obj) throw new Exception("Only one list type is allowed per row.");
                        listProperty = obj as IEnumerable<object>;
                    }
                    else ChangeFromObject(type, propertyName, obj, cell);
                }
                if (listProperty != null) rowsWithList.Add(new Tuple<IXLRow, IEnumerable<object>>(row, listProperty));
            }

            foreach (var row in rowsWithList)
            {
                ChangeFromList(row.Item1, row.Item2, config.FieldValues);
            }
        }

        workbook.SaveAs(GetOutputPath(config.OutputPath));
    }

    private static void ChangeFromObject(Type type, string propertyName, object obj, IXLCell cell)
    {
        var property = type.GetProperty(propertyName);
        if (property == null) return;

        var newValue = type.GetProperty(propertyName)?.GetValue(obj);
        cell.InsertValue(newValue);
    }

    private static void ChangeFromList(IXLRow row, IEnumerable<object> list, Dictionary<string, object> fieldValues)
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
            var newRow = row.CopyTo(row.RowBelow(i));
            InsertValuesInRow(newRow, type, objInList);
        }
        InsertValuesInRow(row, type, firstObject);
    }

    private static void InsertValuesInRow(IXLRow newRow, Type type, object objInList)
    {
        foreach (var cell in newRow.Cells())
        {
            InsertValuesInCell(cell, type, objInList);
        }
    }

    private static void InsertValuesInCell(IXLCell cell, Type type, object objInList)
    {
        var oldValue = cell.GetValue<string>();
        if (!oldValue.StartsWith("{{") && !oldValue.EndsWith("}}")) return;

        oldValue = oldValue.TrimStart('{').TrimEnd('}');

        var fieldInfo = oldValue.Split("::");
        if (fieldInfo.Length != 2) return;

        var propertyName = fieldInfo.ElementAt(1);

        ChangeFromObject(type, propertyName, objInList, cell);
    }

    private static string GetOutputPath(string configOutputPath) => string.IsNullOrEmpty(configOutputPath) ? "./output.xlsx" : configOutputPath;
}