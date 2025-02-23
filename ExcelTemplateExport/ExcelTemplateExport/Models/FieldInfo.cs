namespace ExcelTemplateExport
{
    public class FieldInfo
    {
        public FieldInfo(string value)
        {
            if (value.StartsWith("{{") && value.EndsWith("}}"))
            {
                value = value.TrimStart('{').TrimEnd('}');

                var fieldInfo = value.Split("::");
                ObjectName = fieldInfo.ElementAt(0);
                PropertyName = fieldInfo.ElementAt(1);
            }
        }
        
        public string ObjectName { get; set; }

        public string PropertyName { get; set; }
    }
}