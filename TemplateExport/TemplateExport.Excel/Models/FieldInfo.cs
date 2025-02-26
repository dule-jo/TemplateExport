namespace TemplateExport.Excel.Models
{
    internal class FieldInfo
    {
        internal FieldInfo(string value, ExportConfiguration config)
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
        
        internal string ObjectName { get; set; }

        internal string PropertyName { get; set; }

        internal AggregationType? Aggregation { get; set; }
        
        internal enum AggregationType
        {
            Sum,
            Average,
            Count
        }
    }
}