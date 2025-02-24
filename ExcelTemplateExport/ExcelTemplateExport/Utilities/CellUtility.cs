using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTemplateExport.Utilities
{
    public static class CellUtility
    {
        public static void InsertValue(this IXLCell cell, object? value)
        {
            if (value == null) return;

            if (value is string strValue)
            {
                cell.SetValue(strValue);
            }
            else if (value is int intValue)
            {
                cell.SetValue(intValue);
            }
            else if (value is double doubleValue)
            {
                cell.SetValue(doubleValue);
            }
            else if (value is DateTime dateTimeValue)
            {
                cell.SetValue(dateTimeValue);
            }
            else
            {
                cell.SetValue(value.ToString());
            }
        }

        public static void InsertValuesInRow(IXLRow newRow, Type type, object objInList)
        {
            foreach (var cell in newRow.Cells())
            {
                InsertValuesInCell(cell, type, objInList);
            }
        }

        public static void InsertValuesInRow(IXLRow newRow, Type type, object objInList, IXLRow row)
        {
            foreach (var cell in row.Cells())
            {
                var newCell = newRow.Cell(cell.Address.ColumnNumber);
                InsertValuesInCell(newCell, cell, type, objInList);
            }
        }

        public static void InsertValuesInCell(IXLCell newCell, IXLCell cell, Type type, object objInList)
        {
            var oldValue = cell.GetValue<string>();
        
            var fieldInfo = new FieldInfo(oldValue);
        
            ChangeFromObject(type, fieldInfo, objInList, newCell);
        }

        public static void InsertValuesInCell(IXLCell cell, Type type, object objInList)
        {
            var oldValue = cell.GetValue<string>();
        
            var fieldInfo = new FieldInfo(oldValue);

            ChangeFromObject(type, fieldInfo, objInList, cell);
        }

        public static void ChangeFromObject(Type type, FieldInfo fieldInfo, object obj, IXLCell cell)
        {
            if (fieldInfo.PropertyName == null || fieldInfo.Aggregation != null)
            {
                cell.InsertValue(obj ?? string.Empty);
            }
            else
            {
                var property = type.GetProperty(fieldInfo.PropertyName);
                if (property == null) return;

                var newValue = type.GetProperty(fieldInfo.PropertyName)?.GetValue(obj);
                cell.InsertValue(newValue);
            }
        }
    }
}