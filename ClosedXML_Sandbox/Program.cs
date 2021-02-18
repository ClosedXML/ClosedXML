using ClosedXML.Excel;
using System;
using System.Linq;

namespace ClosedXmlPerformanceTest
{
    class Program
    {
        static string Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        static Random random = new Random();

        static void Main(string[] args)
        {
            TestStyles();
        }

        static void TestStyles()
        {
            const int columnCount = 60;
            const int rowCount = 10000;

            var workbook = new XLWorkbook(XLEventTracking.Disabled);
            var worksheet = workbook.AddWorksheet();
            GenerateData(worksheet, rowCount, columnCount);

            var watch = System.Diagnostics.Stopwatch.StartNew();
            var colorKey = new XLColorKey
            {
                ColorType = XLColorType.Indexed,
                Indexed = 12
            };

            for (var column = 1; column <= columnCount; column++)
            {
                for (var row = 1; row <= rowCount; row++)
                {
                    var cell = worksheet.Cell(row, column);

                    var styles = cell.Style;
                    
                    styles.Modify(s =>
                    {
                        var font = s.Font;
                        font.Bold = true;
                        font.FontSize = 13;
                        font.FontColor = colorKey;

                        var alignment = s.Alignment;
                        alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        alignment.Vertical = XLAlignmentVerticalValues.Justify;

                        var border = s.Border;
                        border.BottomBorder = XLBorderStyleValues.Medium;
                        border.LeftBorder = XLBorderStyleValues.Medium;

                        s.Font = font;
                        s.Alignment = alignment;
                        s.Border = border;

                        return s;
                    });
                }
            }
            watch.Stop();
            Console.WriteLine($"Setting styles elapsed {watch.ElapsedMilliseconds}ms.");
        }

        private static void GenerateData(IXLWorksheet worksheet, int rowCount, int columnCount)
        {
            for (var column = 1; column <= columnCount; column++)
            {
                for (var row = 1; row <= rowCount; row++)
                {
                    var cell = worksheet.Cell(row, column);
                    cell.SetValue(GetRandomTextValue(16));
                }
            }
        }

        private static string GetRandomTextValue(int maxLength)
        {
            return new string(Enumerable.Repeat(Chars, maxLength)
                .Select(s => s[random.Next(s.Length)])
                .ToArray());
        }
    }
}
