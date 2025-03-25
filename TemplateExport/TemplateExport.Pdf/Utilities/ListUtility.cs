using TemplateExport.Pdf.Models;

namespace TemplateExport.Pdf.Utilities;

internal static class ListUtility
{
    internal static double? GetAggregationValue(this IEnumerable<object> list, FieldInfo fieldInfo)
    {
        if (fieldInfo.Aggregation == null || list == null) return null;

        return fieldInfo.Aggregation switch
        {
            FieldInfo.AggregationType.Sum => list.Sum(x => Convert.ToDouble(x.GetType()?.GetProperty(fieldInfo.PropertyName)?.GetValue(x) ?? 0)),
            FieldInfo.AggregationType.Average => list.Where(x=>x.GetType()?.GetProperty(fieldInfo.PropertyName)?.GetValue(x) != null).Average(x => Convert.ToDouble(x.GetType()?.GetProperty(fieldInfo.PropertyName)?.GetValue(x) ?? 0)),
            FieldInfo.AggregationType.Count => list.Count(),
            _ => null
        };
    }
}