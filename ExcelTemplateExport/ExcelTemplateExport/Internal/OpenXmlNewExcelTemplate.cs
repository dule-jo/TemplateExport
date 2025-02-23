using System.Collections;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTemplateExport.Utilities;

namespace ExcelTemplateExport
{
    public class OpenXmlNewExcelTemplate
    {
        public void Export(ExportConfiguration config)
        {
            using var sourceDocument = SpreadsheetDocument.Open(config.TemplatePath, false);
            using var destinationDocument = SpreadsheetDocument.Create(config.OutputPath, SpreadsheetDocumentType.Workbook);

            var sourceWorkbookPart = sourceDocument.WorkbookPart;
            var destinationWorkbookPart = destinationDocument.AddWorkbookPart();
            destinationWorkbookPart.Workbook = new Workbook();
            var destinationSheets = destinationWorkbookPart.Workbook.AppendChild(new Sheets());

            if (sourceWorkbookPart.WorkbookStylesPart != null)
            {
                var sourceStylesPart = sourceWorkbookPart.WorkbookStylesPart;
                var destinationStylesPart = destinationWorkbookPart.AddNewPart<WorkbookStylesPart>();
                destinationStylesPart.Stylesheet = (Stylesheet)sourceStylesPart.Stylesheet.Clone();
                destinationStylesPart.Stylesheet.Save();
            }

            foreach (var sourceSheet in sourceWorkbookPart.Workbook.Sheets.OfType<Sheet>())
            {
                CopyWorksheet(config, sourceWorkbookPart, sourceSheet, destinationWorkbookPart, destinationSheets);
            }

            destinationWorkbookPart.Workbook.Save();
        }

        private void CopyWorksheet(ExportConfiguration config, WorkbookPart sourceWorkbookPart, Sheet sourceSheet, WorkbookPart destinationWorkbookPart, Sheets destinationSheets)
        {
            var sourceWorksheetPart = CreateWorksheet(sourceWorkbookPart, sourceSheet, destinationWorkbookPart, destinationSheets, out var destinationWorksheetPart);

            var sourceSheetData = sourceWorksheetPart.Worksheet.GetFirstChild<SheetData>();
            var destinationSheetData = destinationWorksheetPart.Worksheet.GetFirstChild<SheetData>();

            var rows = sourceWorksheetPart.Worksheet.Descendants<Row>();
            var rowsInserted = 0;
            var newRows = new List<Row>();
            foreach (var row in rows)
            {
                rowsInserted = CopyRow(config, sourceWorkbookPart, row, destinationSheetData, rowsInserted, newRows);
            }
            foreach (var newRow in newRows)
            {
                destinationSheetData.Append(newRow);
            }
        }

        private static int CopyRow(ExportConfiguration config, WorkbookPart sourceWorkbookPart, Row row, SheetData? destinationSheetData, int rowsInserted, List<Row> newRows)
        {
            var cells = row.Descendants<Cell>();

            var listObject = IsListRow(cells, config, sourceWorkbookPart);

            if (listObject != null)
            {
                foreach (var member in listObject)
                {
                    foreach (var cell in cells)
                    {
                        CopyCell(config, cell, destinationSheetData, sourceWorkbookPart, rowsInserted, member, newRows);
                    }
                    if (member != listObject.LastOrDefault())
                        rowsInserted++;
                }
            }
            else
            {
                foreach (var cell in cells)
                {
                    CopyCell(config, cell, destinationSheetData, sourceWorkbookPart, rowsInserted, null, newRows);
                }
            }
            return rowsInserted;
        }

        private static IEnumerable<object> IsListRow(IEnumerable<Cell> cells, ExportConfiguration config, WorkbookPart sourceWorkbookPart)
        {
            IEnumerable<object> response = null;
            foreach (var cell in cells)
            {
                var oldValue = GetCellValue(cell, sourceWorkbookPart);
                var fieldInfo = new FieldInfo(oldValue);
                if (fieldInfo.ObjectName == null || fieldInfo.PropertyName == null) continue;
                if (!config.FieldValues.TryGetValue(fieldInfo.ObjectName, out var obj)) continue;
                var type = obj.GetType();
                if (typeof(IEnumerable).IsAssignableFrom(type))
                {
                    if (obj is IEnumerable<object> && response != null && response != obj as IEnumerable<object>)
                    {
                        throw new Exception("Only one list type is allowed per row.");
                    }

                    response = obj as IEnumerable<object>;
                }
            }
            return response;
        }

        private static void CopyCell(ExportConfiguration config, Cell cell, SheetData? destinationSheetData, WorkbookPart sourceWorkbookPart, int rowsInserted, object listObject, List<Row> newRows)
        {
            var rowIndex = ((IConvertible)cell.CellReference.Value.Substring(1)).ToUInt32(null);
            var destinationRow = destinationSheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex == rowIndex);

            if (destinationRow == null)
            {
                destinationRow = new Row { RowIndex = rowIndex };
                newRows.Add(destinationRow);
            }

            var oldValue = GetCellValue(cell, sourceWorkbookPart);

            var fieldInfo = new FieldInfo(oldValue);
            if (fieldInfo.ObjectName == null || fieldInfo.PropertyName == null)
            {
                CopyValueToDestination(cell, oldValue, destinationRow, rowsInserted);
                return;
            }

            if (!config.FieldValues.TryGetValue(fieldInfo.ObjectName, out var obj)) return;
            var type = obj.GetType();

            if (typeof(IEnumerable).IsAssignableFrom(type))
            {
                if (listObject == null) return;
                var memberType = listObject.GetType();
                var value = memberType.GetProperty(fieldInfo.PropertyName)?.GetValue(listObject);
                CopyValueToDestination(cell, value?.ToString() ?? string.Empty, destinationRow, rowsInserted);
            }
            else
            {
                var value = type.GetProperty(fieldInfo.PropertyName)?.GetValue(obj);
                CopyValueToDestination(cell, value?.ToString() ?? string.Empty, destinationRow, rowsInserted);
            }
        }

        private static void CopyValueToDestination(Cell cell, string oldValue, Row destinationRow, int rowsInserted)
        {
            var cellReference = cell.CellReference;
            if (rowsInserted > 0)
            {
                var rowNumber = ((IConvertible)cellReference.Value.Substring(1)).ToUInt32(null);
                cellReference = cellReference.Value.Replace(rowNumber.ToString(), (rowNumber + rowsInserted).ToString());
            }
            var destinationCell = new Cell
            {
                CellReference = cellReference,
                DataType = cell.DataType,
                StyleIndex = cell.StyleIndex
            };

            destinationCell.InsertValue(oldValue);
            destinationRow.Append(destinationCell);
        }

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.CellValue == null)
                return string.Empty;

            var value = cell.CellValue.Text;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTablePart = workbookPart.SharedStringTablePart;
                if (stringTablePart != null)
                    value = stringTablePart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }

            return value;
        }

        private static WorksheetPart CreateWorksheet(WorkbookPart sourceWorkbookPart, Sheet sourceSheet, WorkbookPart destinationWorkbookPart, Sheets destinationSheets, out WorksheetPart destinationWorksheetPart)
        {
            var sourceWorksheetPart = (WorksheetPart)sourceWorkbookPart.GetPartById(sourceSheet.Id);
            destinationWorksheetPart = destinationWorkbookPart.AddNewPart<WorksheetPart>();
            destinationWorksheetPart.Worksheet = new Worksheet(sourceWorksheetPart.Worksheet.OuterXml);

            var sheetId = destinationSheets.Elements<Sheet>().Select(s => (uint)s.SheetId.Value).DefaultIfEmpty(0u).Max() + 1;
            var destinationSheet = new Sheet
            {
                Id = destinationWorkbookPart.GetIdOfPart(destinationWorksheetPart),
                SheetId = sheetId,
                Name = sourceSheet.Name
            };
            destinationSheets.Append(destinationSheet);
            return sourceWorksheetPart;
        }
    }
}