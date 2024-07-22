using OpenXmlPractice.Core;

var path = @"G:\Solutions\my-openxml-practice\Examples";
var reports = new[]
{
    "Report1.docx",
    "Report2.docx"
};
var dataBase = "Database.xlsx";

var nameTag = "username";
var placeTag = "movewher";

foreach (var report in reports)
{
    var reportPath = Path.Combine(path, report);
    var dataBasePath = Path.Combine(path, dataBase);

    var name = DocxReader.ReadControllValue(reportPath, nameTag);
    var place = DocxReader.ReadControllValue(reportPath, placeTag);

    var rowId = XlsxReader.GetRow(dataBasePath, "users", "A", name);
    var cellReference = XlsxReader.GetCellReference(dataBasePath, "users", 1, "Place");

    XlsxReader.UpdateCellValue(dataBasePath, "users", rowId, cellReference, place);
}
