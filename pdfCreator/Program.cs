// See https://aka.ms/new-console-template for more information
using ExcelReader;
using pdfSharp;

const string xlsFileName = "testData.xls";
const int rowCount = 20;
const int colCount = 7;

// generate xls file with random strings 
ExcelWorker.GenerateRandomExcelFile(xlsFileName, rowCount, colCount);

//read xls
var array = ExcelWorker.ReadExcelFile(xlsFileName);

//print table to pdf
PdfPrinter.Print(array);