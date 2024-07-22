using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using OpenXmlPractice.Core.Extensions;

namespace OpenXmlPractice.Core;

public static class XlsxReader
{
    public static string FindFirstEmptyCellInColumn(
        string filePath,
        string sheetName,
        string columnName)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false);

        var workbookPart = document.WorkbookPart;
        var worksheet = workbookPart.GetWorksheetByName(sheetName);

        var rows = worksheet.Descendants<Row>();

        foreach (var row in worksheet.Descendants<Row>())
        {
            var cellReference = columnName + row.RowIndex;

            var cell = row
                .Elements<Cell>()
                .FirstOrDefault(c => c.CellReference == cellReference);

            if (cell == null || string.IsNullOrEmpty(workbookPart.GetCellValue(cell)))
            {
                return cellReference;
            }
        }

        throw new InvalidOperationException("No empty cell found in the column");
    }
    public static string GetCellReference(
        string filePath,
        string sheetName,
        uint rowIndex,
        string value)
    {
        using var document = SpreadsheetDocument.Open(filePath, false);

        var workbookPart = document.WorkbookPart;
        var row = workbookPart
            .GetWorksheetByName(sheetName)
            .GetRowByIndex(rowIndex);

        var cellReference = string.Empty;

        foreach (var cell in row.Elements<Cell>())
        {
            if (workbookPart.GetCellValue(cell) == value)
            {
                cellReference = cell.CellReference.Value;
                break;
            }
        }

        if (string.IsNullOrEmpty(cellReference))
        {
            return null;
        }

        return Regex.Replace(cellReference, @"[\d-]", string.Empty);
    }

    public static void UpdateCellValue(
        string filePath,
        string sheetName,
        uint rowIndex,
        string cellReference,
        string newValue)
    {
        using var  document = SpreadsheetDocument.Open(filePath, true);

        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

        if (sheet == null)
        {
            Console.WriteLine("Sheet not found.");
            return;
        }

        var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
        var worksheet = worksheetPart.Worksheet;

        // Get the specified row
        var row = worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            throw new ArgumentException("The specified row does not exist.");
        }

        var a = row.Elements<Cell>()
            .Select(_ => _.CellReference);

        // Get the cell in the specified column and row
        Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == $"{cellReference}{rowIndex}");
        if (cell == null)
        {
            // If the cell does not exist, create a new one
            cell = new Cell() { CellReference = cellReference };
            row.Append(cell);
        }

        // Update the cell value
        cell.CellValue = new CellValue(newValue);
        cell.DataType = new EnumValue<CellValues>(CellValues.String);

        worksheetPart.Worksheet.Save();
    }

    public static uint GetRow(
        string filePath,
        string sheetName,
        string cellReference,
        string userName)
    {
        using var document = SpreadsheetDocument.Open(filePath, true);

        var workbookPart = document.WorkbookPart;

        var worksheet = workbookPart.GetWorksheetByName(sheetName);

        // Iterate through the rows to find the one with the specified value in the first column
        Row targetRow = null;
        var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();

        foreach (var row in rows)
        {
            var firstCell = row
                .Elements<Cell>()
                .FirstOrDefault(c => c.CellReference.Value.StartsWith(cellReference));

            if (firstCell != null && workbookPart.GetCellValue(firstCell) == userName)
            {
                targetRow = row;
                break;
            }
        }

        if (targetRow == null)
        {
            Console.WriteLine("The specified value was not found in the first column.");
            return 0;
        }

        return targetRow.RowIndex;
    }
}
