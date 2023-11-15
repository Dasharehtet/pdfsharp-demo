using OfficeOpenXml;
using System.Text;

namespace ExcelReader;

public static class ExcelWorker
{
    public static string[,] ReadExcelFile(string fileName)
    {
        ExcelPackage.LicenseContext = LicenseContext.Commercial;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage(new FileInfo(fileName));
        // Get the first worksheet
        var worksheet = package.Workbook.Worksheets[0];

        // Get the dimensions of the worksheet
        var start = worksheet.Dimension.Start;
        var end = worksheet.Dimension.End;

        var rowCount = end.Row;
        var colCount = end.Column;

        var array = new string[colCount, rowCount];

        // Read data from the worksheet
        for (var row = start.Row; row <= end.Row; row++)
        {
            for (var col = start.Column; col <= end.Column; col++)
            {
                // Access cell value by row and column
                var cellValue = worksheet.Cells[row, col].Text;
                    
                // Print or process the cell value as needed
                array[col-1, row-1] = cellValue;
            }
        }
            
        return array;
    }
    
    public static void GenerateRandomExcelFile(string fileName, int rowCount, int colCount)
    {
       // ExcelPackage.LicenseContext = LicenseContext.Commercial;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage();
        // Create a new worksheet
        var worksheet = package.Workbook.Worksheets.Add("Sheet1");

        // Generate and fill random data
        for (var row = 1; row <= rowCount; row++)
        {
            for (var col = 1; col <= colCount; col++)
            {
                worksheet.Cells[row, col].Value = GenerateRandomString();
            }
        }

        // Save the Excel file
        var excelFile = new FileInfo(fileName);
        package.SaveAs(excelFile);
    }

    private static string GenerateRandomString()
    {
        var random = new Random();
        var length = random.Next(3, 51); // Random length between 3 and 50

        var stringBuilder = new StringBuilder();

        for (var i = 0; i < length;)
        {
            var wordLength = random.Next(3, 6); // Random word length between 3 and 5
            var randomWord = GenerateRandomWord(wordLength);

            stringBuilder.Append(randomWord);

            i += wordLength;

            // Add space between words if there's still space in the string
            if (i < length)
            {
                stringBuilder.Append(" ");
                i++; // Increment for the space character
            }
        }

        return stringBuilder.ToString();
    }

    private static string GenerateRandomWord(int length)
    {
        const string characters = "abcdefghijklmnopqrstuvwxyz";

        var random = new Random();
        var wordBuilder = new StringBuilder();

        for (var i = 0; i < length; i++)
        {
            var randomChar = characters[random.Next(characters.Length)];
            wordBuilder.Append(randomChar);
        }

        return wordBuilder.ToString();
    }
}