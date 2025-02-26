using System.Collections;
using HtmlAgilityPack;
using TemplateExport.Pdf.Models;

namespace TemplateExport.Pdf.Internal;

public class TemplateExportPdf : ITemplateExportPdf
{
    public void Export(PdfExportConfiguration config)
    {
        var templateHtml = new HtmlDocument();
        templateHtml.Load(config.TemplatePath);

        TraverseNode(templateHtml.DocumentNode, config);

        templateHtml.Save(config.OutputPath);
    }

    private void TraverseNode(HtmlNode node, PdfExportConfiguration config)
    {
        var removeNodes = new List<HtmlNode>();
        var insertAfterNodes = new List<Tuple<HtmlNode, List<HtmlNode>>>();
        foreach (var childNode in node.ChildNodes)
        {
            if (childNode.Attributes.Any(x => x.Name == config.IfAttribute)) removeNodes.Add(childNode);
            if (childNode.Attributes.Any(x => x.Name == config.ForAttribute)) ReplaceWithList(config, removeNodes, childNode, insertAfterNodes);
            else TraverseNode(childNode, config);
        }

        ReplaceListNodes(insertAfterNodes);

        RemoveNodes(removeNodes);
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
        foreach (var insertAfterNode in insertAfterNodes)
        {
            var deleteNode = insertAfterNode.Item1;
            var newNodes = insertAfterNode.Item2;
            var parent = deleteNode.ParentNode;

            foreach (var obj in newNodes)
            {
                parent.InsertAfter(obj, deleteNode);
            }

            deleteNode.Remove();
        }
    }

    private static void ReplaceWithList(PdfExportConfiguration config, List<HtmlNode> removeNodes, HtmlNode childNode, List<Tuple<HtmlNode, List<HtmlNode>>> insertAfterNodes)
    {
        removeNodes.Add(childNode);

        var setName = childNode.GetAttributeValue(config.ForAttribute, "");
        if (!config.DataSets.TryGetValue(setName, out var list)) return;
        var type = list?.GetType() ?? typeof(string);

        childNode.Attributes.Remove(config.ForAttribute);

        if (typeof(IEnumerable).IsAssignableFrom(type) && type != typeof(string))
        {
            var newList = new List<HtmlNode>();
            foreach (var obj in list as IEnumerable<object>)
            {
                newList.Add(CreateNewNode(childNode, obj, setName));
            }

            insertAfterNodes.Add(new Tuple<HtmlNode, List<HtmlNode>>(childNode, newList));
        }
    }

    private static HtmlNode CreateNewNode(HtmlNode childNode, object o, string setName)
    {
        return HtmlNode.CreateNode(childNode.OuterHtml);
    }
}