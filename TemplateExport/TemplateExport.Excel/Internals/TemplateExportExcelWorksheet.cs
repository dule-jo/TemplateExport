using System.Collections;
using ClosedXML.Excel;
using ExcelTemplateExport.Models;
using ExcelTemplateExport.Utilities;

namespace ExcelTemplateExport.Internals
{
    internal class TemplateExportExcelWorksheet
    {
        private ExcelExportConfiguration _config;
        private IXLStyle _defaultStyle = null;
        private int _rowsInserted = 0;
        private readonly List<object[]> _dataList = new();
        private readonly Dictionary<(int row, int col), IXLStyle> _originalCellStyles = new();
        private readonly Dictionary<(int row, int col), IXLStyle> _cellStyles = new();
        private readonly Dictionary<(int row, int col), (int numOfRows, int numOfCols)> _originalMergedCells = new();
        private readonly Dictionary<(int row, int col), (int numOfRows, int numOfCols)> _mergedCells = new();

        internal void Export(ExcelExportConfiguration config, IXLWorksheet templateSheet, IXLWorksheet outputSheet)
        {
            _config = config;

            var lastCol = templateSheet.LastColumnUsed().ColumnNumber();

            var lastVisitedRow = 1;
            foreach (var row in templateSheet.Rows())
            {
                AddSkippedRowsToDataSet(row, lastVisitedRow, lastCol);
                
                lastVisitedRow = row.RowNumber();
                var rowValues = new object[lastCol];
                IEnumerable<object> listObject = null;
                
                foreach (var cell in row.Cells())
                {
                    if (cell.IsEmpty() || cell.DataType == XLDataType.Blank) continue;
                    
                    if (config.PreserveCellStyles) SaveStyle(row.RowNumber(), cell.Address.ColumnNumber, cell);
                    if (config.PreserveMergeCells) SaveMergedCell(cell, row.RowNumber(), cell.Address.ColumnNumber);
                                       
                    AddValueToData(cell, rowValues, ref listObject);
                }
                
                if (listObject == null)
                {
                    _dataList.Add(rowValues);
                }
                else
                {
                    CopyList(listObject, rowValues, templateSheet, row.RowNumber());
                }
            }

            outputSheet.Cell(1, 1).InsertData(_dataList);

            MergeCells(outputSheet);

            CopyCellsStyle(outputSheet);

            SetRowHeight(config, templateSheet, outputSheet);

            SetColumnWidth(config, lastCol, templateSheet, outputSheet);
        }

        private void AddValueToData(IXLCell cell, object[] rowValues, ref IEnumerable<object>? listObject)
        {
            var text = GetTextInside(cell);
            var fieldInfo = new FieldInfo(text, _config);
            if (!TryGetType(fieldInfo, out var obj, out var type))
            {
                rowValues[cell.Address.ColumnNumber - 1] = text;
                return;
            }
            if (fieldInfo.Aggregation != null)
            {
                rowValues[cell.Address.ColumnNumber - 1] = (obj as IEnumerable<object>).GetAggregationValue(fieldInfo);
            }
            else if (typeof(IEnumerable).IsAssignableFrom(type) && type != typeof(string))
            {
                var list = obj as IEnumerable<object>;
                if (list == null) return;

                listObject = list;
                rowValues[cell.Address.ColumnNumber - 1] = text;
            }
            else
            {
                var newValue = obj;

                if (fieldInfo.PropertyName != null)
                {
                    var property = type.GetProperty(fieldInfo.PropertyName);
                    if (property == null) return;

                    newValue = type.GetProperty(fieldInfo.PropertyName)?.GetValue(obj);
                }

                rowValues[cell.Address.ColumnNumber - 1] = newValue;
            }
        }

        private void AddSkippedRowsToDataSet(IXLRow row, int lastVisitedRow, int lastCol)
        {
            for (var i = 0; i < row.RowNumber() - lastVisitedRow - 1; i++)
                _dataList.Add(new object[lastCol]);
        }

        private void SaveMergedCell(IXLCell cell, int r, int c)
        {
            if (cell.IsMerged() && cell.DataType != XLDataType.Blank) 
                _mergedCells[(r + _rowsInserted, c)] = (cell.MergedRange().RowCount(), cell.MergedRange().ColumnCount());
        }

        private void SaveStyle(int r, int c, IXLCell cell)
        {
            _cellStyles[(r + _rowsInserted, c)] = cell.Style;
        }

        private void CopyList(IEnumerable<object> listObject, object[] rowValues, IXLWorksheet templateSheet, int r)
        {
            var i = 0;
            foreach (var member in listObject)
            {
                _rowsInserted++;
                var copy = new object[rowValues.Length];
                Array.Copy(rowValues, copy, rowValues.Length);
                var index = 0;
                foreach (var value in copy)
                {
                    if (i == 0)
                    {
                        var cell = templateSheet.Cell(r, index + 1);
                        if (cell.IsMerged())
                        {
                            _originalMergedCells[(r, index + 1)] = (cell.MergedRange().RowCount(), cell.MergedRange().ColumnCount());
                            _mergedCells[(r + _rowsInserted, index + 1)] = (cell.MergedRange().RowCount(), cell.MergedRange().ColumnCount());
                        }
                        _originalCellStyles[(r, index + 1)] = cell.Style;
                    }
                    
                    {
                        _cellStyles[(r + _rowsInserted - 1, index + 1)] = _originalCellStyles[(r, index + 1)];

                        if (_originalMergedCells.ContainsKey((r, index + 1)))
                        {
                            _mergedCells[(r + _rowsInserted - 1, index + 1)] = _originalMergedCells[(r, index + 1)];
                        }
                    }
                    
                    var fieldInfo = new FieldInfo(value?.ToString(), _config);
                    if (fieldInfo.ObjectName == null) continue;
                    var property = member.GetType().GetProperty(fieldInfo.PropertyName);
                    var newValue = property?.GetValue(member);
                    copy[index++] = newValue;
                }

                _dataList.Add(copy);
            }
            i++;
            _rowsInserted--; // to compensate for row in template
        }

        private void MergeCells(IXLWorksheet outputSheet)
        {
            if (!_config.PreserveMergeCells) return;
            
            foreach (var mergedCell in _mergedCells)
            {
                var startCell = outputSheet.Cell(mergedCell.Key.row, mergedCell.Key.col);
                var endCell = outputSheet.Cell(mergedCell.Key.row + mergedCell.Value.numOfRows - 1, mergedCell.Key.col + mergedCell.Value.numOfCols - 1);
                var range = outputSheet.Range(startCell, endCell);
                range.Merge();
            }
        }

        private void CopyCellsStyle(IXLWorksheet outputSheet)
        {
            if (!_config.PreserveCellStyles) return;
            
            foreach (var key in _cellStyles.Keys)
            {
                if (!_cellStyles.TryGetValue((key.row, key.col), out var style)) continue;

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
        }

        private void SetRowHeight(ExcelExportConfiguration config, IXLWorksheet templateSheet, IXLWorksheet outputSheet)
        {
            if (config.AutoFitRows) outputSheet.Rows().AdjustToContents();
            if (!config.PreserveRowHeight) return;
            
            foreach (var r in _cellStyles.Keys.Select(x => x.row).Distinct())
            {
                var templateRow = templateSheet.Row(r);
                var outputRow = outputSheet.Row(r);

                outputRow.Height = templateRow.Height;
            }
        }

        private void SetColumnWidth(ExcelExportConfiguration config, int lastCol, IXLWorksheet templateSheet, IXLWorksheet outputSheet)
        {
            if (config.AutoFitColumns) outputSheet.Columns().AdjustToContents();
            if (!config.PreserveColumnWidth) return;
            
            for (var c = 1; c <= lastCol; c++)
            {
                var templateColumn = templateSheet.Column(c);
                var outputColumn = outputSheet.Column(c);

                outputColumn.Width = templateColumn.Width;
            }
        }

        private static string GetTextInside(IXLCell cell)
        {
            if (cell.HasFormula) throw new Exception("Export does not support cells with formulas");
            return cell.IsEmpty() ? string.Empty : cell.GetText();
        }

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

}