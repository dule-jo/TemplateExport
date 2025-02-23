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

        public static void InsertValue(this Cell cell, object value)
        {
            switch (value)
            {
                case string strValue:
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(strValue);
                    break;
                case int intValue:
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(intValue.ToString());
                    break;
                case double doubleValue:
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(doubleValue.ToString());
                    break;
                case DateTime dateTimeValue:
                    cell.DataType = CellValues.Date;
                    cell.CellValue = new CellValue(dateTimeValue.ToString("yyyy-MM-ddTHH:mm:ss.fff"));
                    break;
                default:
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(value.ToString());
                    break;
            }
        }

        public static string GetCellValue(SharedStringItem[] values, Cell x) => x.CellValue != null ? values[int.Parse(x.CellValue.Text)].InnerText : string.Empty;

        public static bool IsMeantForReplace(string x) => x.StartsWith("{{") && x.EndsWith("}}") && x.Contains("::");
    }
}