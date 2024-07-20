using DocumentFormat.OpenXml;
using System.Buffers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace OpenXmlPractice.Core;

public static class XlsxReader
{
    public static void UpdateCellValue(string filePath, string sheetName, uint rowIndex, string cellReference, string newValue)
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

        var sheet = workbookPart
            .Workbook
            .Descendants<Sheet>()
            .FirstOrDefault(s => s.Name == sheetName);

        if (sheet == null)
        {
            Console.WriteLine("Sheet not found.");
            return 0;
        }

        var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
        var worksheet = worksheetPart.Worksheet;

        // Iterate through the rows to find the one with the specified value in the first column
        Row targetRow = null;
        var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();

        foreach (var row in rows)
        {
            Cell firstCell = row
                .Elements<Cell>()
                .FirstOrDefault(c => c.CellReference.Value.StartsWith(cellReference));

            if (firstCell != null && GetCellValue(document, firstCell) == userName)
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

    public static string GetCellReference(
        string filePath,
        string sheetName,
        uint rowIndex,
        string searchValue)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false);

        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart
            .Workbook
            .Descendants<Sheet>()
            .FirstOrDefault(s => s.Name == sheetName);

        if (sheet == null)
        {
            Console.WriteLine("Sheet not found.");
            return null;
        }

        var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
        var worksheet = worksheetPart.Worksheet;

        // Get the specified row
        var row = worksheet
            .GetFirstChild<SheetData>()
            .Elements<Row>()
            .FirstOrDefault(r => r.RowIndex == rowIndex);

        if (row == null)
        {
            Console.WriteLine("The specified row does not exist.");
            return null;
        }

        // Iterate through the cells in the row to find the cell with the specified value
        foreach (Cell cell in row.Elements<Cell>())
        {
            if (GetCellValue(document, cell) == searchValue)
            {
                return Regex.Replace(cell.CellReference.Value, @"[\d-]", string.Empty);
            }
        }

        // Return null if the value is not found
        return null;
    }



    static string GetCellValue(
        SpreadsheetDocument document,
        Cell cell)
    {
        if (cell.CellValue == null)
        {
            return null;
        }

        string value = cell.CellValue.Text;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            return document.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
        }

        return value;
    }

    static Cell GetCell(
        Row row,
        string columnName)
    {
        return row.Elements<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, columnName + row.RowIndex, true) == 0);
    }
}
