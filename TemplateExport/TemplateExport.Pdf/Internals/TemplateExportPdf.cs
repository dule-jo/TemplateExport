using System.Collections;
using HtmlAgilityPack;
using TemplateExport.Pdf.Models;
using TemplateExport.Pdf.Utilities;

namespace TemplateExport.Pdf.Internals;

public class TemplateExportPdf : ITemplateExportPdf
{
    private PdfExportConfiguration _config;

    public void Export(PdfExportConfiguration config)
    {
        _config = config;
        var templateHtml = new HtmlDocument();
        templateHtml.Load(_config.TemplatePath);

        TraverseNode(templateHtml.DocumentNode);

        templateHtml.Save(_config.OutputPath);
    }

    private void TraverseNode(HtmlNode node)
    {
        if (ReplaceIf(node)) return;
        
        foreach (var childNode in node.ChildNodes)
        {
            TraverseNode(childNode);
        }

        if (node.Attributes.Any(x => x.Name == _config.ForAttribute)) ReplaceWithList(node);
        else
        {
            var fieldInfos = PatternUtility.ExtractPatterns(node.InnerHtml, _config);

            foreach (var fieldInfo in fieldInfos)
            {
                ReplaceInnerHtml(node, fieldInfo);
            }
        }
    }

    private bool ReplaceIf(HtmlNode node)
    {
        if (node.Attributes.Any(x => x.Name == _config.IfAttribute))
        {
            node.Remove();
            return true;
        }
        return false;
    }

    private void ReplaceInnerHtml(HtmlNode node, FieldInfo fieldInfo)
    {
        object newValue2 = string.Empty;
        if (TryGetType(fieldInfo, out var obj, out var type))
        {
            if (fieldInfo.Aggregation != null)
            {
                newValue2 = (obj as IEnumerable<object>).GetAggregationValue(fieldInfo);
            }
            else if (typeof(IEnumerable).IsAssignableFrom(type) && type != typeof(string))
            {
            }
            else
            {
                var property = type.GetProperty(fieldInfo.PropertyName);
                if (property == null) return;

                newValue2 = type.GetProperty(fieldInfo.PropertyName)?.GetValue(obj);
            }
        }

        node.InnerHtml = node.InnerHtml.Replace(fieldInfo.Value, newValue2.ToString());
    }

    private static void RemoveNodes(List<HtmlNode> removeNodes)
    {
        foreach (var removeNode in removeNodes)
        {
            removeNode.Remove();
        }
    }

    private static void ReplaceListNodes(List<Tuple<HtmlNode, List<HtmlNode>>> insertAfterNodes)
    {
        foreach (var (deleteNode, newNodes) in insertAfterNodes)
        {
            var parent = deleteNode.ParentNode;

            foreach (var obj in newNodes)
            {
                parent.InsertAfter(obj, deleteNode);
            }

            deleteNode.Remove();
        }
    }

    private void ReplaceWithList(HtmlNode childNode)
    {
        // removeNodes.Add(childNode);
        //
        // var setName = childNode.GetAttributeValue(_config.ForAttribute, "");
        // if (!_config.DataSets.TryGetValue(setName, out var list)) return;
        // var type = list?.GetType() ?? typeof(string);
        //
        // childNode.Attributes.Remove(_config.ForAttribute);
        //
        // if (typeof(IEnumerable).IsAssignableFrom(type) && type != typeof(string))
        // {
        //     var newList = new List<HtmlNode>();
        //     var i = 0;
        //     foreach (var obj in list as IEnumerable<object>)
        //     {
        //         var htmlNode = CreateNewNode(childNode, obj, setName);
        //
        //         htmlNode.InnerHtml = htmlNode.InnerHtml.Replace($"", "");
        //
        //         newList.Add(htmlNode);
        //     }
        //
        //     insertAfterNodes.Add(new Tuple<HtmlNode, List<HtmlNode>>(childNode, newList));
        // }
        //
        // foreach (var node in insertAfterNodes.SelectMany(x => x.Item2))
        // {
        //     TraverseNode(node);
        // }
    }

    private static HtmlNode CreateNewNode(HtmlNode childNode, object o, string setName) { return HtmlNode.CreateNode(childNode.OuterHtml); }

    private bool TryGetType(FieldInfo fieldInfo, out object? obj, out Type? type)
    {
        if (fieldInfo.ObjectName == null)
        {
            obj = null;
            type = null;
            return false;
        }

        if (!_config.DataSets.TryGetValue(fieldInfo.ObjectName, out obj))
        {
            type = null;
            return false;
        }
        type = obj?.GetType() ?? typeof(string);
        return true;
    }
}