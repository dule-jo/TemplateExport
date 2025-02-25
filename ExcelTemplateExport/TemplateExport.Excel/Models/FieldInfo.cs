namespace ExcelTemplateExport.Models
{
    public class FieldInfo
    {
        public FieldInfo(string value, ExportConfiguration config)
        {
            if (!value.StartsWith(config.TemplateStringStartsWith) || !value.EndsWith(config.TemplateStringEndsWith)) return;
            
            value = value.Substring(config.TemplateStringStartsWith.Length, value.Length - config.TemplateStringStartsWith.Length - config.TemplateStringEndsWith.Length);

            if (!value.Contains(config.TemplateStringSeparator))
            {
                ObjectName = value;
                return;
            }

            var fieldInfo = value.Split(config.TemplateStringSeparator);
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