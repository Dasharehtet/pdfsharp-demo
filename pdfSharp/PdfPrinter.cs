using PdfSharpCore.Drawing;
using PdfSharpCore.Drawing.Layout;
using PdfSharpCore.Pdf;

namespace pdfSharp;

public static class PdfPrinter
{
    private const int FontSize = 12;
    private const string FontFamilyName = "Arial";
    private const string DocumentTitle = "Dynamic Table";
    private const string PdfFileName = "DynamicTableWithBordersAndWrapping.pdf";
    
    public static void Print(string[,] data)
    {
        // Create PDF document
        var document = new PdfDocument();
        document.Info.Title = DocumentTitle;

        // Create a font for the table
        var font = new XFont(FontFamilyName, FontSize, XFontStyle.Regular);

        // Create a page
        var page = document.AddPage();
        var gfx = XGraphics.FromPdfPage(page);

        // Determine the number of columns based on the data array
        var colCount = data.GetLength(0);
        var columnWidth = page.Width.Point / colCount;  

        // Draw the table with borders and text wrapping
        PdfDrawer.DrawTable(gfx, font, columnWidth, data, document);

        // Save the PDF to a file
        document.Save(PdfFileName);
    }
}