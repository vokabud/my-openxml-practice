// See https://aka.ms/new-console-template for more information
using OpenXmlPractice.Console;

Console.WriteLine("Hello, World!");


var filePath = @"G:\Solutions\my-openxml-practice\Examples";
var fileName = "Report.docx";

DocxReader.Read(filePath, fileName);

