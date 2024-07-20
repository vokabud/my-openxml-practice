using OpenXmlPractice.Core;

var report = @"G:\Solutions\my-openxml-practice\Examples\Report.docx";
var dataBase = @"G:\Solutions\my-openxml-practice\Examples\Database.xlsx";

var nameTag = "username";
var placeTag = "movewher";

var name = DocxReader.ReadControllValue(report, nameTag);
var place = DocxReader.ReadControllValue(report, placeTag);

Console.WriteLine($"{name} => {place}");

var rowId = XlsxReader.GetRow(dataBase, "users", "A", name);
var cellReference = XlsxReader.GetCellReference(dataBase, "users", 1, "Place");

Console.WriteLine(rowId);
Console.WriteLine(cellReference);

XlsxReader.UpdateCellValue(dataBase, "users", rowId, cellReference, place);
