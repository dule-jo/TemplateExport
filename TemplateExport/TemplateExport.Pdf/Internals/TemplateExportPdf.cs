using System.Collections;
using HtmlAgilityPack;
using iText.Html2pdf;
using TemplateExport.Pdf.Models;
using TemplateExport.Pdf.Utilities;

namespace TemplateExport.Pdf.Internals;

internal class TemplateExportPdf : ITemplateExportPdf
{
    private PdfExportConfiguration _config;

    public void Export(PdfExportConfiguration config)
    {
        _config = config;
        var templateHtml = GetTemplateHtml();

        TraverseNode(templateHtml.DocumentNode);
        
        SaveToOutput(templateHtml);
    }

    private void SaveToOutput(HtmlDocument templateHtml)
    {
        if (_config.OutputPath != null) ConvertHtmlToPdfFile(templateHtml.DocumentNode.OuterHtml, _config.OutputPath);
        else if (_config.OutputStream != null) HtmlConverter.ConvertToPdf(templateHtml.DocumentNode.OuterHtml, _config.OutputStream);
        else throw new Exception("OutputPath or OutputStream must be set.");
    }

    private HtmlDocument GetTemplateHtml()
    {
        var templateHtml = new HtmlDocument();

        if (_config.TemplatePath == null) return LoadHtmlFromHeadAndBody();

        templateHtml.Load(_config.TemplatePath);
        return templateHtml;

    }

    private HtmlDocument LoadHtmlFromHeadAndBody()
    {
        var doc = new HtmlDocument();
        doc.LoadHtml("<html><head></head><body></body></html>");

        var header = doc.DocumentNode.SelectSingleNode("//head");
        foreach (var headFile in _config.TemplateHead)
        {
            var headDoc = new HtmlDocument();
            headDoc.Load(headFile);

            foreach (var headNode in headDoc.DocumentNode.ChildNodes)
            {
                header.AppendChild(headNode);
            }
        }

        var body = doc.DocumentNode.SelectSingleNode("//body");
        foreach (var bodyFile in _config.TemplateBody)
        {
            var bodyDoc = new HtmlDocument();
            bodyDoc.Load(bodyFile);

            foreach (var bodyNode in bodyDoc.DocumentNode.ChildNodes)
            {
                body.AppendChild(bodyNode);
            }
        }

        return doc;
    }

    private void TraverseNode(HtmlNode node, Dictionary<string, object> dataSetsForList = null)
    {
        if (ReplaceIf(node, dataSetsForList)) return;
        if (ReplaceElse(node)) return;

        foreach (var childNode in node.ChildNodes.ToList())
        {
            TraverseNode(childNode, dataSetsForList);
        }

        if (node.Attributes.Any(x => x.Name == _config.ForAttribute)) ReplaceWithList(node);
        else
        {
            var fieldInfos = PatternUtility.ExtractPatterns(node.InnerHtml, _config);

            foreach (var fieldInfo in fieldInfos)
            {
                ReplaceInnerHtml(node, fieldInfo, dataSetsForList);
            }
        }
    }

    private bool ReplaceElse(HtmlNode node)
    {
        if (node.Attributes.All(x => x.Name != _config.ElseAttribute)) return false;
        node.Remove();
        return true;
    }

    private bool ReplaceIf(HtmlNode node, Dictionary<string, object> dataSetsForList)
    {
        if (node.Attributes.All(x => x.Name != _config.IfAttribute)) return false;

        var nodeAttribute = node.Attributes[_config.IfAttribute];
        var ifAttributeValue = nodeAttribute.Value;

        nodeAttribute.Remove();

        var fieldInfo = new FieldInfo(ifAttributeValue, _config);

        if (!TryGetType(fieldInfo, dataSetsForList, out var obj, out var type) || typeof(IEnumerable).IsAssignableFrom(type) && type != typeof(string)) return false;

        if (fieldInfo.Aggregation != null) return false;

        var property = type.GetProperty(fieldInfo.PropertyName);
        if (property == null) return false;

        if (GetBoolValue(type.GetProperty(fieldInfo.PropertyName)?.GetValue(obj))) return false;

        return IfRemove(node);
    }

    private bool IfRemove(HtmlNode node)
    {
        var nextNode = GetNextRealSibling(node);
        if (nextNode != null)
        {
            var elseAttribute = nextNode.Attributes.Any(x => x.Name == _config.ElseAttribute) ? nextNode.Attributes[_config.ElseAttribute] : null;
            elseAttribute?.Remove();
        }
        node.Remove();
        return true;
    }

    private static HtmlNode GetNextRealSibling(HtmlNode node)
    {
        var nextSibling = node.NextSibling;
        while (nextSibling != null && nextSibling.NodeType != HtmlNodeType.Element)
        {
            nextSibling = nextSibling.NextSibling;
        }
        return nextSibling;
    }

    private static bool GetBoolValue(object? getValue)
    {
        return getValue switch
        {
            bool b => b,
            int i => i != 0,
            double d => d != 0,
            string s => !string.IsNullOrEmpty(s),
            _ => false
        };
    }

    private void ReplaceInnerHtml(HtmlNode node, FieldInfo fieldInfo, Dictionary<string, object> dataSetsForList)
    {
        object newValue2 = string.Empty;
        if (!TryGetType(fieldInfo, dataSetsForList, out var obj, out var type) || typeof(IEnumerable).IsAssignableFrom(type) && type != typeof(string)) return;

        if (fieldInfo.Aggregation != null)
        {
            newValue2 = (obj as IEnumerable<object>).GetAggregationValue(fieldInfo);
        }
        else
        {
            var property = type.GetProperty(fieldInfo.PropertyName);
            if (property == null) return;

            newValue2 = type.GetProperty(fieldInfo.PropertyName)?.GetValue(obj);
        }

        node.InnerHtml = node.InnerHtml.Replace(fieldInfo.Value, newValue2.ToString());
    }

    private void ReplaceWithList(HtmlNode childNode)
    {
        var setName = childNode.GetAttributeValue(_config.ForAttribute, "");
        
        var fieldInfo = new FieldInfo(setName, _config);
        if (!TryGetType(fieldInfo, null, out var list, out var type)) return;
        
        childNode.Attributes.Remove(_config.ForAttribute);

        if (typeof(IEnumerable).IsAssignableFrom(type) && type != typeof(string))
        {
            var i = 0;
            foreach (var obj in list as IEnumerable<object>)
            {
                var htmlNode = CreateNewNode(childNode, obj, fieldInfo.ObjectName);
                childNode.ParentNode.InsertAfter(htmlNode, childNode);
                TraverseNode(htmlNode, new Dictionary<string, object> { { fieldInfo.ObjectName, obj } });
            }
        }

        childNode.Remove();
    }

    private static HtmlNode CreateNewNode(HtmlNode childNode, object o, string setName) { return HtmlNode.CreateNode(childNode.OuterHtml); }

    private bool TryGetType(FieldInfo fieldInfo, Dictionary<string, object> dataSetsForList, out object? obj, out Type? type)
    {
        if (fieldInfo.ObjectName == null)
        {
            obj = null;
            type = null;
            return false;
        }

        if (dataSetsForList != null && dataSetsForList.TryGetValue(fieldInfo.ObjectName, out obj))
        {
            type = obj?.GetType() ?? typeof(string);
            return true;
        }

        if (!_config.DataSets.TryGetValue(fieldInfo.ObjectName, out obj))
        {
            type = null;
            return false;
        }

        type = obj?.GetType() ?? typeof(string);
        return true;
    }

    private static void ConvertHtmlToPdfFile(string html, string outputPath)
    {
        using var stream = new FileStream(outputPath, FileMode.Create);
        HtmlConverter.ConvertToPdf(html, stream);
    }
}