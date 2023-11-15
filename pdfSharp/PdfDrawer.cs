using PdfSharpCore.Drawing;
using PdfSharpCore.Drawing.Layout;
using PdfSharpCore.Pdf;

namespace pdfSharp;

public static class PdfDrawer
{
    private static double _fontHeight;
    private static PdfDocument _document = null!;
    private static int _rowCount;
    private static int _colCount;
    private static double _columnWidth;
    private static XPoint _currentPosition;
    private static readonly XStringFormat Format = new()
    {
        LineAlignment = XLineAlignment.Near,
        Alignment = XStringAlignment.Near
    };
    
    public static void DrawTable(XGraphics gfx, XFont font, double columnWidth, string[,] data, PdfDocument document)
    {
        _document = document;
        
        _colCount = data.GetLength(0);
        _rowCount = data.GetLength(1);
        _columnWidth = columnWidth;
        
        _fontHeight = font.GetHeight();
        
        // Set border thickness
        var borderPen = new XPen(XColors.Black, 1);

        // Calculate the total width of the table
        var totalWidth = gfx.PageSize.Width;

        // Draw the table borders
        gfx.DrawRectangle(borderPen, 0, 0, totalWidth, _fontHeight);

        // Draw the table header
        DrawTableHeader(gfx, font, borderPen);

        // Draw the table rows with text wrapping
        DrawTableRows(gfx, font, data, borderPen);
    }

    private static void DrawTableHeader(XGraphics gfx, XFont font, XPen borderPen)
    {
        _currentPosition = new XPoint(0, 0);
        
        for (var i = 0; i < _colCount; i++)
        {
            // Draw cell border
            gfx.DrawRectangle(borderPen, _currentPosition.X, _currentPosition.Y, _columnWidth, _fontHeight);

            // Draw cell content
            gfx.DrawString($"Column {i + 1}", font, XBrushes.Black, _currentPosition.X + 3, _currentPosition.Y + font.Size);

            _currentPosition.X += _columnWidth;
        }
        
        // Move to the next line after drawing the header
        _currentPosition.Y += _fontHeight;
        
        // Reset X position for the new row
        _currentPosition.X = 0;
    }

    private static void DrawTableRows(XGraphics gfx, XFont font, string[,] data, XPen borderPen)
    {
        var textFormatter = new XTextFormatter(gfx);

        for (var row = 1; row <= _rowCount; row++)
        {
            var rowHeight = RowHeightCalculation(gfx, font, data, row);
            
            gfx = CreateNewPage(gfx, font, borderPen, rowHeight, row, ref textFormatter);

            PrintText(font, data, row, textFormatter, rowHeight);

            PrintRectangles(gfx, borderPen, rowHeight);
        }
    }

    private static void PrintRectangles(XGraphics gfx, XPen borderPen, double rowHeight)
    {
        for (var col = 0; col < _colCount; col++)
        {
            // Draw cell border
            gfx.DrawRectangle(borderPen, _currentPosition.X, _currentPosition.Y, _columnWidth, rowHeight);

            // Move to the next column
            _currentPosition.X += _columnWidth;
        }

        // Move to the next row
        _currentPosition.Y += rowHeight;

        // Reset X position for the new row
        _currentPosition.X = 0;
    }

    private static void PrintText(XFont font, string[,] data, int row, XTextFormatter textFormatter, double rowHeight)
    {
        for (var col = 0; col < _colCount; col++)
        {
            // Get the data for the cell
            var cellData = data[col, row - 1];

            // Draw cell content with text wrapping
            textFormatter.DrawString(cellData, font, XBrushes.Black,
                new XRect(_currentPosition.X + 2, _currentPosition.Y + 2, _columnWidth - 4, rowHeight), Format);

            // Move to the next column
            _currentPosition.X += _columnWidth;
        }

        _currentPosition.X = 0;
    }

    private static XGraphics CreateNewPage(XGraphics gfx, XFont font, XPen borderPen, double rowHeight, int row, ref XTextFormatter textFormatter)
    {
        // Check for page break and create a new page if needed
        if (_currentPosition.Y + rowHeight > gfx.PageSize.Height - 30 && row < _rowCount - 1)
        {
            // Create a new page
            var page = _document.AddPage();
            gfx = XGraphics.FromPdfPage(page);
            textFormatter = new XTextFormatter(gfx);

            // Redraw the table header on the new page
            DrawTableHeader(gfx, font, borderPen);
        }

        return gfx;
    }

    private static double RowHeightCalculation(XGraphics gfx, XFont font, string[,] data, int row)
    {
        var rowHeight = _fontHeight;
        
        for (var col = 0; col < _colCount; col++)
        {
            // Get the data for the cell
            var cellData = data[col, row - 1];

            // Calculate row height
            var textWidth = gfx.MeasureString(cellData, font).Width;
            var lineCount = (int) Math.Ceiling(textWidth / _columnWidth);

            if (lineCount * _fontHeight > rowHeight)
            {
                rowHeight = (lineCount * _fontHeight) + 2;
            }
        }

        return rowHeight;
    }
}