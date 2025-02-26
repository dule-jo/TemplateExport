using TemplateExport.Excel.Models;

namespace TemplateExport.Excel.Utilities;

internal static class ListUtility
{
    internal static double? GetAggregationValue(this IEnumerable<object> list, FieldInfo fieldInfo)
    {
        if (fieldInfo.Aggregation == null || list == null) return null;

        return fieldInfo.Aggregation switch
        {
            FieldInfo.AggregationType.Sum => list.Sum(x => Convert.ToDouble(x.GetType().GetProperty(fieldInfo.PropertyName).GetValue(x))),
            FieldInfo.AggregationType.Average => list.Average(x => Convert.ToDouble(x.GetType().GetProperty(fieldInfo.PropertyName).GetValue(x))),
            FieldInfo.AggregationType.Count => list.Count(),
            _ => null
        };
    }
}