using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlPractice.Core.Exceptions;

namespace OpenXmlPractice.Core.Extensions;

public static class WorksheetExtensions
{
    public static Row GetRowByIndex(
        this Worksheet worksheet,
        uint rowIndex)
    {
        var row = worksheet
            .GetFirstChild<SheetData>()
            .Elements<Row>()
            .FirstOrDefault(r => r.RowIndex == rowIndex);

        return row == null 
            ? throw new RowNotFoundException()
            : row;
    }
}
