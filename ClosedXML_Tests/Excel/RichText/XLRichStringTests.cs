using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    /// <summary>
    ///     This is a test class for XLRichStringTests and is intended
    ///     to contain all XLRichStringTests Unit Tests
    /// </summary>
    [TestFixture]
    public class XLRichStringTests
    {
        [Test]
        public void AccessRichTextTest1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.RichText.AddText("12");
            cell.DataType = XLCellValues.Number;

            Assert.AreEqual(12.0, cell.GetDouble());

            IXLRichText richText = cell.RichText;

            Assert.AreEqual("12", richText.ToString());

            richText.AddText("34");

            Assert.AreEqual("1234", cell.GetString());

            Assert.AreEqual(XLCellValues.Number, cell.DataType);

            Assert.AreEqual(1234.0, cell.GetDouble());
        }

        /// <summary>
        ///     A test for AddText
        /// </summary>
        [Test]
        public void AddTextTest1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            IXLRichText richString = cell.RichText;

            string text = "Hello";
            richString.AddText(text).SetBold().SetFontColor(XLColor.Red);

            Assert.AreEqual(cell.GetString(), text);
            Assert.AreEqual(cell.RichText.First().Bold, true);
            Assert.AreEqual(cell.RichText.First().FontColor, XLColor.Red);

            Assert.AreEqual(1, richString.Count);

            richString.AddText("World");
            Assert.AreEqual(richString.First().Text, text, "Item in collection is not the same as the one returned");
        }

        [Test]
        public void AddTextTest2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            Int32 number = 123;

            cell.SetValue(number).Style
                .Font.SetBold()
                .Font.SetFontColor(XLColor.Red);

            string text = number.ToString();

            Assert.AreEqual(cell.RichText.ToString(), text);
            Assert.AreEqual(cell.RichText.First().Bold, true);
            Assert.AreEqual(cell.RichText.First().FontColor, XLColor.Red);

            Assert.AreEqual(1, cell.RichText.Count);

            cell.RichText.AddText("World");
            Assert.AreEqual(cell.RichText.First().Text, text, "Item in collection is not the same as the one returned");
        }

        [Test]
        public void AddTextTest3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            Int32 number = 123;
            cell.Value = number;
            cell.Style
                .Font.SetBold()
                .Font.SetFontColor(XLColor.Red);

            string text = number.ToString();

            Assert.AreEqual(cell.RichText.ToString(), text);
            Assert.AreEqual(cell.RichText.First().Bold, true);
            Assert.AreEqual(cell.RichText.First().FontColor, XLColor.Red);

            Assert.AreEqual(1, cell.RichText.Count);

            cell.RichText.AddText("World");
            Assert.AreEqual(cell.RichText.First().Text, text, "Item in collection is not the same as the one returned");
        }

        /// <summary>
        ///     A test for Clear
        /// </summary>
        [Test]
        public void ClearTest()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World!");

            richString.ClearText();
            String expected = String.Empty;
            String actual = richString.ToString();
            Assert.AreEqual(expected, actual);

            Assert.AreEqual(0, richString.Count);
        }

        [Test]
        public void CountTest()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World!");

            Assert.AreEqual(3, richString.Count);
        }

        [Test]
        public void HasRichTextTest1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.RichText.AddText("123");

            Assert.AreEqual(true, cell.HasRichText);

            cell.DataType = XLCellValues.Text;

            Assert.AreEqual(true, cell.HasRichText);

            cell.DataType = XLCellValues.Number;

            Assert.AreEqual(false, cell.HasRichText);

            cell.RichText.AddText("123");

            Assert.AreEqual(true, cell.HasRichText);

            cell.Value = 123;

            Assert.AreEqual(false, cell.HasRichText);

            cell.RichText.AddText("123");

            Assert.AreEqual(true, cell.HasRichText);

            cell.SetValue("123");

            Assert.AreEqual(false, cell.HasRichText);
        }

        /// <summary>
        ///     A test for Characters
        /// </summary>
        [Test]
        public void Substring_All_From_OneString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");

            IXLFormattedText<IXLRichText> actual = richString.Substring(0);

            Assert.AreEqual(richString.First(), actual.First());

            Assert.AreEqual(1, actual.Count);

            actual.First().SetBold();

            Assert.AreEqual(true, ws.Cell(1, 1).RichText.First().Bold);
        }

        [Test]
        public void Substring_All_From_ThreeStrings()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(0);

            Assert.AreEqual(richString.ElementAt(0), actual.ElementAt(0));
            Assert.AreEqual(richString.ElementAt(1), actual.ElementAt(1));
            Assert.AreEqual(richString.ElementAt(2), actual.ElementAt(2));

            Assert.AreEqual(3, actual.Count);
            Assert.AreEqual(3, richString.Count);

            actual.First().SetBold();

            Assert.AreEqual(true, ws.Cell(1, 1).RichText.First().Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.Last().Bold);
        }

        [Test]
        public void Substring_From_OneString_End()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");

            IXLFormattedText<IXLRichText> actual = richString.Substring(2);

            Assert.AreEqual(1, actual.Count); // substring was in one piece

            Assert.AreEqual(2, richString.Count); // The text was split because of the substring

            Assert.AreEqual("llo", actual.First().Text);

            Assert.AreEqual("He", richString.First().Text);
            Assert.AreEqual("llo", richString.Last().Text);

            actual.First().SetBold();

            Assert.AreEqual(false, ws.Cell(1, 1).RichText.First().Bold);
            Assert.AreEqual(true, ws.Cell(1, 1).RichText.Last().Bold);

            richString.Last().SetItalic();

            Assert.AreEqual(false, ws.Cell(1, 1).RichText.First().Italic);
            Assert.AreEqual(true, ws.Cell(1, 1).RichText.Last().Italic);

            Assert.AreEqual(true, actual.First().Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).RichText.First().FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.Last().FontSize);

            Assert.AreEqual(20, actual.First().FontSize);
        }

        [Test]
        public void Substring_From_OneString_Middle()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");

            IXLFormattedText<IXLRichText> actual = richString.Substring(2, 2);

            Assert.AreEqual(1, actual.Count); // substring was in one piece

            Assert.AreEqual(3, richString.Count); // The text was split because of the substring

            Assert.AreEqual("ll", actual.First().Text);

            Assert.AreEqual("He", richString.First().Text);
            Assert.AreEqual("ll", richString.ElementAt(1).Text);
            Assert.AreEqual("o", richString.Last().Text);

            actual.First().SetBold();

            Assert.AreEqual(false, ws.Cell(1, 1).RichText.First().Bold);
            Assert.AreEqual(true, ws.Cell(1, 1).RichText.ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.Last().Bold);

            richString.Last().SetItalic();

            Assert.AreEqual(false, ws.Cell(1, 1).RichText.First().Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(1).Italic);
            Assert.AreEqual(true, ws.Cell(1, 1).RichText.Last().Italic);

            Assert.AreEqual(false, actual.First().Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).RichText.First().FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(1).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.Last().FontSize);

            Assert.AreEqual(20, actual.First().FontSize);
        }

        [Test]
        public void Substring_From_OneString_Start()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");

            IXLFormattedText<IXLRichText> actual = richString.Substring(0, 2);

            Assert.AreEqual(1, actual.Count); // substring was in one piece

            Assert.AreEqual(2, richString.Count); // The text was split because of the substring

            Assert.AreEqual("He", actual.First().Text);

            Assert.AreEqual("He", richString.First().Text);
            Assert.AreEqual("llo", richString.Last().Text);

            actual.First().SetBold();

            Assert.AreEqual(true, ws.Cell(1, 1).RichText.First().Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.Last().Bold);

            richString.Last().SetItalic();

            Assert.AreEqual(false, ws.Cell(1, 1).RichText.First().Italic);
            Assert.AreEqual(true, ws.Cell(1, 1).RichText.Last().Italic);

            Assert.AreEqual(false, actual.First().Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).RichText.First().FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.Last().FontSize);

            Assert.AreEqual(20, actual.First().FontSize);
        }

        [Test]
        public void Substring_From_ThreeStrings_End1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(21);

            Assert.AreEqual(1, actual.Count); // substring was in one piece

            Assert.AreEqual(4, richString.Count); // The text was split because of the substring

            Assert.AreEqual("bors!", actual.First().Text);

            Assert.AreEqual("Good Morning", richString.ElementAt(0).Text);
            Assert.AreEqual(" my ", richString.ElementAt(1).Text);
            Assert.AreEqual("neigh", richString.ElementAt(2).Text);
            Assert.AreEqual("bors!", richString.ElementAt(3).Text);

            actual.First().SetBold();

            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(0).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(2).Bold);
            Assert.AreEqual(true, ws.Cell(1, 1).RichText.ElementAt(3).Bold);

            richString.Last().SetItalic();

            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(0).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(1).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(2).Italic);
            Assert.AreEqual(true, ws.Cell(1, 1).RichText.ElementAt(3).Italic);

            Assert.AreEqual(true, actual.First().Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(0).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(1).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(2).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(3).FontSize);

            Assert.AreEqual(20, actual.First().FontSize);
        }

        [Test]
        public void Substring_From_ThreeStrings_End2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(13);

            Assert.AreEqual(2, actual.Count);

            Assert.AreEqual(4, richString.Count); // The text was split because of the substring

            Assert.AreEqual("my ", actual.ElementAt(0).Text);
            Assert.AreEqual("neighbors!", actual.ElementAt(1).Text);

            Assert.AreEqual("Good Morning", richString.ElementAt(0).Text);
            Assert.AreEqual(" ", richString.ElementAt(1).Text);
            Assert.AreEqual("my ", richString.ElementAt(2).Text);
            Assert.AreEqual("neighbors!", richString.ElementAt(3).Text);

            actual.ElementAt(1).SetBold();

            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(0).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(2).Bold);
            Assert.AreEqual(true, ws.Cell(1, 1).RichText.ElementAt(3).Bold);

            richString.Last().SetItalic();

            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(0).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(1).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(2).Italic);
            Assert.AreEqual(true, ws.Cell(1, 1).RichText.ElementAt(3).Italic);

            Assert.AreEqual(false, actual.ElementAt(0).Italic);
            Assert.AreEqual(true, actual.ElementAt(1).Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(0).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(1).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(2).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(3).FontSize);

            Assert.AreEqual(20, actual.ElementAt(0).FontSize);
            Assert.AreEqual(20, actual.ElementAt(1).FontSize);
        }

        [Test]
        public void Substring_From_ThreeStrings_Mid1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(5, 10);

            Assert.AreEqual(2, actual.Count);

            Assert.AreEqual(5, richString.Count); // The text was split because of the substring

            Assert.AreEqual("Morning", actual.ElementAt(0).Text);
            Assert.AreEqual(" my", actual.ElementAt(1).Text);

            Assert.AreEqual("Good ", richString.ElementAt(0).Text);
            Assert.AreEqual("Morning", richString.ElementAt(1).Text);
            Assert.AreEqual(" my", richString.ElementAt(2).Text);
            Assert.AreEqual(" ", richString.ElementAt(3).Text);
            Assert.AreEqual("neighbors!", richString.ElementAt(4).Text);
        }

        [Test]
        public void Substring_From_ThreeStrings_Mid2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(5, 15);

            Assert.AreEqual(3, actual.Count);

            Assert.AreEqual(5, richString.Count); // The text was split because of the substring

            Assert.AreEqual("Morning", actual.ElementAt(0).Text);
            Assert.AreEqual(" my ", actual.ElementAt(1).Text);
            Assert.AreEqual("neig", actual.ElementAt(2).Text);

            Assert.AreEqual("Good ", richString.ElementAt(0).Text);
            Assert.AreEqual("Morning", richString.ElementAt(1).Text);
            Assert.AreEqual(" my ", richString.ElementAt(2).Text);
            Assert.AreEqual("neig", richString.ElementAt(3).Text);
            Assert.AreEqual("hbors!", richString.ElementAt(4).Text);
        }

        [Test]
        public void Substring_From_ThreeStrings_Start1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(0, 4);

            Assert.AreEqual(1, actual.Count); // substring was in one piece

            Assert.AreEqual(4, richString.Count); // The text was split because of the substring

            Assert.AreEqual("Good", actual.First().Text);

            Assert.AreEqual("Good", richString.ElementAt(0).Text);
            Assert.AreEqual(" Morning", richString.ElementAt(1).Text);
            Assert.AreEqual(" my ", richString.ElementAt(2).Text);
            Assert.AreEqual("neighbors!", richString.ElementAt(3).Text);

            actual.First().SetBold();

            Assert.AreEqual(true, ws.Cell(1, 1).RichText.ElementAt(0).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(2).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(3).Bold);

            richString.First().SetItalic();

            Assert.AreEqual(true, ws.Cell(1, 1).RichText.ElementAt(0).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(1).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(2).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(3).Italic);

            Assert.AreEqual(true, actual.First().Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(0).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(1).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(2).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(3).FontSize);

            Assert.AreEqual(20, actual.First().FontSize);
        }

        [Test]
        public void Substring_From_ThreeStrings_Start2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(0, 15);

            Assert.AreEqual(2, actual.Count);

            Assert.AreEqual(4, richString.Count); // The text was split because of the substring

            Assert.AreEqual("Good Morning", actual.ElementAt(0).Text);
            Assert.AreEqual(" my", actual.ElementAt(1).Text);

            Assert.AreEqual("Good Morning", richString.ElementAt(0).Text);
            Assert.AreEqual(" my", richString.ElementAt(1).Text);
            Assert.AreEqual(" ", richString.ElementAt(2).Text);
            Assert.AreEqual("neighbors!", richString.ElementAt(3).Text);

            actual.ElementAt(1).SetBold();

            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(0).Bold);
            Assert.AreEqual(true, ws.Cell(1, 1).RichText.ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(2).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(3).Bold);

            richString.First().SetItalic();

            Assert.AreEqual(true, ws.Cell(1, 1).RichText.ElementAt(0).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(1).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(2).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).RichText.ElementAt(3).Italic);

            Assert.AreEqual(true, actual.ElementAt(0).Italic);
            Assert.AreEqual(false, actual.ElementAt(1).Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(0).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(1).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(2).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).RichText.ElementAt(3).FontSize);

            Assert.AreEqual(20, actual.ElementAt(0).FontSize);
            Assert.AreEqual(20, actual.ElementAt(1).FontSize);
        }

        [Test]
        public void Substring_IndexOutsideRange1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");

            Assert.That(() => richString.Substring(50), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");
            richString.AddText("World");

            Assert.That(() => richString.Substring(50), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");

            Assert.That(() => richString.Substring(1, 10), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange4()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");
            richString.AddText("World");

            Assert.That(() => richString.Substring(5, 20), Throws.TypeOf<IndexOutOfRangeException>());
        }

        /// <summary>
        ///     A test for ToString
        /// </summary>
        [Test]
        public void ToStringTest()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World");
            string expected = "Hello World";
            string actual = richString.ToString();
            Assert.AreEqual(expected, actual);

            richString.AddText("!");
            expected = "Hello World!";
            actual = richString.ToString();
            Assert.AreEqual(expected, actual);

            richString.ClearText();
            expected = String.Empty;
            actual = richString.ToString();
            Assert.AreEqual(expected, actual);
        }
    }
}