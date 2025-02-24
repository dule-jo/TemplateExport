namespace ExcelTemplateExport.Utilities;

public static class ListUtility
{
    public static double? GetAggregationValue(this IEnumerable<object> list, FieldInfo fieldInfo)
    {
        if (fieldInfo.Aggregation == null || list == null) return null;
        
        if (fieldInfo.Aggregation == FieldInfo.AggregationType.Sum)
            return list.Sum(x => Convert.ToDouble(x.GetType().GetProperty(fieldInfo.PropertyName).GetValue(x)));
        if (fieldInfo.Aggregation == FieldInfo.AggregationType.Average)
            return list.Average(x => Convert.ToDouble(x.GetType().GetProperty(fieldInfo.PropertyName).GetValue(x)));
        if (fieldInfo.Aggregation == FieldInfo.AggregationType.Count)
            return list.Count();
        
        return null;
    }
}