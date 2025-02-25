using System.Collections;
using ClosedXML.Excel;
using ExcelTemplateExport.Utilities;

namespace ExcelTemplateExport
{
    public class NewClosedXmlExcelTemplate
    {
        private ExportConfiguration _config;
        private static int rowsInserted = 0;
        private Dictionary<string, object[,]> sheetData = new();

        public void Export(ExportConfiguration config)
        {
            _config = config;

            using var templateWb = new XLWorkbook(_config.TemplatePath);
            using var outputWb = new XLWorkbook();
            var templateSheet = templateWb.Worksheet(1);
            var outputSheet = outputWb.AddWorksheet("Output");

            var lastCol = templateSheet.LastColumnUsed().ColumnNumber();

            var dataList = new List<object[]>();
            var originalCellStyles = new Dictionary<(int row, int col), IXLStyle>();
            var cellStyles = new Dictionary<(int row, int col), IXLStyle>();
            var originalMergedCells = new Dictionary<(int row, int col), (int numOfRows, int numOfCols)>();
            var mergedCells = new Dictionary<(int row, int col), (int numOfRows, int numOfCols)>();

            var lastVisitedRow = 1;
            foreach (var row in templateSheet.Rows())
            {
                var r = row.RowNumber();
                for (var i = 0; i < r - lastVisitedRow - 1; i++)
                    dataList.Add(new object[lastCol]);
                lastVisitedRow = r;
                var rowValues = new object[lastCol];

                IEnumerable<object> listObject = null;
                foreach (var cell in row.Cells())
                {
                    if (cell.IsEmpty() || cell.DataType == XLDataType.Blank) continue;
                    var c = cell.Address.ColumnNumber;
                    cellStyles[(r + rowsInserted, c)] = cell.Style;
                    if (cell.IsMerged() && cell.DataType != XLDataType.Blank) mergedCells[(r + rowsInserted, c)] = (cell.MergedRange().RowCount(), cell.MergedRange().ColumnCount());
                    var text = GetTextInside(cell);
                    var fieldInfo = new FieldInfo(text);
                    if (fieldInfo.ObjectName == null)
                    {
                        rowValues[c - 1] = text;
                        continue;
                    }
                    else
                    {
                        if (!TryGetType(fieldInfo, out var obj, out var type)) continue;
                        if (fieldInfo.Aggregation != null)
                        {
                            var aggregateValue = (obj as IEnumerable<object>).GetAggregationValue(fieldInfo);
                            rowValues[c - 1] = aggregateValue;
                        }
                        else if (typeof(IEnumerable).IsAssignableFrom(type) && type != typeof(string))
                        {
                            var list = obj as IEnumerable<object>;
                            if (list == null) continue;

                            listObject = list;
                            rowValues[c - 1] = text;
                        }
                        else
                        {
                            var property = type.GetProperty(fieldInfo.PropertyName);
                            if (property == null) continue;

                            var newValue = type.GetProperty(fieldInfo.PropertyName)?.GetValue(obj);
                            rowValues[c - 1] = newValue;
                        }
                    }
                }
                if (listObject == null)
                {
                    dataList.Add(rowValues);
                }
                else
                {
                    foreach (var member in listObject)
                    {
                        rowsInserted++;
                        var copy = new object[rowValues.Length];
                        Array.Copy(rowValues, copy, rowValues.Length);
                        var index = 0;
                        foreach (var value in copy)
                        {
                            if (member == listObject.FirstOrDefault())
                            {
                                var cell = templateSheet.Cell(r, index + 1);
                                if (cell.IsMerged())
                                {
                                    originalMergedCells[(r, index + 1)] = (cell.MergedRange().RowCount(), cell.MergedRange().ColumnCount());
                                    mergedCells[(r + rowsInserted, index + 1)] = (cell.MergedRange().RowCount(), cell.MergedRange().ColumnCount());
                                }
                                originalCellStyles[(r, index + 1)] = cell.Style;
                            }
                            else
                            {
                                cellStyles[(r + rowsInserted - 1, index + 1)] = originalCellStyles[(r, index + 1)];

                                if (originalMergedCells.ContainsKey((r, index + 1)))
                                {
                                    mergedCells[(r + rowsInserted - 1, index + 1)] = originalMergedCells[(r, index + 1)];
                                }
                            }
                            var fieldInfo = new FieldInfo(value?.ToString());
                            if (fieldInfo.ObjectName == null) continue;
                            var property = member.GetType().GetProperty(fieldInfo.PropertyName);
                            var newValue = property?.GetValue(member);
                            copy[index++] = newValue;
                        }

                        dataList.Add(copy);
                    }
                    rowsInserted--; // to compensate for row in template
                }
            }

            outputSheet.Cell(1, 1).InsertData(dataList);

            foreach (var mergedCell in mergedCells)
            {
                var startCell = outputSheet.Cell(mergedCell.Key.row, mergedCell.Key.col);
                var endCell = outputSheet.Cell(mergedCell.Key.row + mergedCell.Value.numOfRows - 1, mergedCell.Key.col + mergedCell.Value.numOfCols - 1);
                var range = outputSheet.Range(startCell, endCell);
                range.Merge();
            }

            foreach (var key in cellStyles.Keys)
            {
                if (!cellStyles.TryGetValue((key.row, key.col), out var style)) continue;

                var outputCell = outputSheet.Cell(key.row, key.col);
                if (outputCell.IsMerged())
                {
                    var range = outputSheet.Range(outputCell, outputCell.MergedRange().LastCell());
                    range.Style = style;
                }
                else
                {
                    outputCell.Style = style;
                }
            }

            if (config.PreserveRowHeight)
            {
                foreach (var r in cellStyles.Keys.Select(x => x.row).Distinct())
                {
                    var templateRow = templateSheet.Row(r);
                    var outputRow = outputSheet.Row(r);

                    outputRow.Height = templateRow.Height;
                }
            }

            if (config.PreserveColumnWidth)
            {
                for (var c = 1; c <= lastCol; c++)
                {
                    var templateColumn = templateSheet.Column(c);
                    var outputColumn = outputSheet.Column(c);

                    outputColumn.Width = templateColumn.Width;
                }
            }

            outputWb.SaveAs(config.OutputPath);

            Console.WriteLine("Podaci i stilovi su uspešno kopirani!");
        }

        private static string GetTextInside(IXLCell cell) { return cell.IsEmpty() ? string.Empty : cell.GetText(); }

        private bool TryGetType(FieldInfo fieldInfo, out object? obj, out Type? type)
        {
            if (!_config.FieldValues.TryGetValue(fieldInfo.ObjectName, out obj))
            {
                type = null;
                return false;
            }
            type = obj?.GetType() ?? typeof(string);
            return true;
        }
    }

}