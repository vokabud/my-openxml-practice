using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlPractice.Core.Exceptions;

namespace OpenXmlPractice.Core.Extensions;

public static class WorkbookPartExtensions
{
    public static Worksheet GetWorksheetByName(
        this WorkbookPart workbookPart,
        string name)
    {
        var sheet = workbookPart
            .Workbook
            .Descendants<Sheet>()
            .FirstOrDefault(s => s.Name == name)
            ?? throw new SheetNotFoundException(name);

        var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));

        return worksheetPart.Worksheet;
    }

    public static string GetCellValue(
        this WorkbookPart workbookPart,
        Cell cell)
    {
        if (cell == null || cell.CellValue == null)
        {
            return string.Empty;
        }

        string value = cell.CellValue.InnerText;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value));
            return ssi.Text.Text;
        }

        return value;
    }
}
