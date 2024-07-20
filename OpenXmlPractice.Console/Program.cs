using OpenXmlPractice.Core;

var filePath = @"G:\Solutions\my-openxml-practice\Examples\Report.docx";

var nameTag = "username";
var placeTag = "movewher";

var name = DocxReader.ReadControllValue(filePath, nameTag);
var place = DocxReader.ReadControllValue(filePath, placeTag);

Console.WriteLine($"{name} => {place}");
