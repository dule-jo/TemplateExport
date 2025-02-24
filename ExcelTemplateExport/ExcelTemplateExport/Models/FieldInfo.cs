namespace ExcelTemplateExport
{
    public class FieldInfo
    {
        public FieldInfo(string value)
        {
            if (!value.StartsWith("{{") || !value.EndsWith("}}")) return;
            
            value = value.TrimStart('{').TrimEnd('}');

            if (!value.Contains("::"))
            {
                ObjectName = value;
                return;
            }

            var fieldInfo = value.Split("::");
            ObjectName = fieldInfo.ElementAt(0);
            PropertyName = fieldInfo.ElementAt(1);
            if (fieldInfo.Length == 3)
            {
                Aggregation = fieldInfo.ElementAt(2) switch
                {
                    "Sum" => AggregationType.Sum,
                    "Average" => AggregationType.Average,
                    "Count" => AggregationType.Count,
                    _ => null
                };
            }
        }
        
        public string ObjectName { get; set; }

        public string PropertyName { get; set; }

        public AggregationType? Aggregation { get; set; }
        
        public enum AggregationType
        {
            Sum,
            Average,
            Count
        }
    }
}