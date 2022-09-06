using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.RichText
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
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.CreateRichText().AddText("12");
            cell.DataType = XLDataType.Number;

            Assert.AreEqual(12.0, cell.GetDouble());

            var richText = cell.GetRichText();

            Assert.AreEqual("12", richText.ToString());

            richText.AddText("34");

            Assert.AreEqual("1234", cell.GetString());

            Assert.AreEqual(XLDataType.Number, cell.DataType);

            Assert.AreEqual(1234.0, cell.GetDouble());
        }

        /// <summary>
        ///     A test for AddText
        /// </summary>
        [Test]
        public void AddTextTest1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var richString = cell.CreateRichText();

            var text = "Hello";
            richString.AddText(text).SetBold().SetFontColor(XLColor.Red);

            Assert.AreEqual(cell.GetString(), text);
            Assert.AreEqual(cell.GetRichText().First().Bold, true);
            Assert.AreEqual(cell.GetRichText().First().FontColor, XLColor.Red);

            Assert.AreEqual(1, richString.Count);

            richString.AddText("World");
            Assert.AreEqual(richString.First().Text, text, "Item in collection is not the same as the one returned");
        }

        [Test]
        public void AddTextTest2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var number = 123;

            cell.SetValue(number).Style
                .Font.SetBold()
                .Font.SetFontColor(XLColor.Red);

            var text = number.ToString();

            Assert.AreEqual(cell.GetRichText().ToString(), text);
            Assert.AreEqual(cell.GetRichText().First().Bold, true);
            Assert.AreEqual(cell.GetRichText().First().FontColor, XLColor.Red);

            Assert.AreEqual(1, cell.GetRichText().Count);

            cell.GetRichText().AddText("World");
            Assert.AreEqual(cell.GetRichText().First().Text, text, "Item in collection is not the same as the one returned");
        }

        [Test]
        public void AddTextTest3()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var number = 123;
            cell.Value = number;
            cell.Style
                .Font.SetBold()
                .Font.SetFontColor(XLColor.Red);

            var text = number.ToString();

            Assert.AreEqual(cell.GetRichText().ToString(), text);
            Assert.AreEqual(cell.GetRichText().First().Bold, true);
            Assert.AreEqual(cell.GetRichText().First().FontColor, XLColor.Red);

            Assert.AreEqual(1, cell.GetRichText().Count);

            cell.GetRichText().AddText("World");
            Assert.AreEqual(cell.GetRichText().First().Text, text, "Item in collection is not the same as the one returned");
        }

        /// <summary>
        ///     A test for Clear
        /// </summary>
        [Test]
        public void ClearTest()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World!");

            richString.ClearText();
            var expected = string.Empty;
            var actual = richString.ToString();
            Assert.AreEqual(expected, actual);

            Assert.AreEqual(0, richString.Count);
        }

        [Test]
        public void CountTest()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World!");

            Assert.AreEqual(3, richString.Count);
        }

        [Test]
        public void HasRichTextTest1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.GetRichText().AddText("123");

            Assert.AreEqual(true, cell.HasRichText);

            cell.DataType = XLDataType.Text;

            Assert.AreEqual(true, cell.HasRichText);

            cell.DataType = XLDataType.Number;

            Assert.AreEqual(false, cell.HasRichText);

            cell.GetRichText().AddText("123");

            Assert.AreEqual(true, cell.HasRichText);

            cell.Value = 123;

            Assert.AreEqual(false, cell.HasRichText);

            cell.GetRichText().AddText("123");

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
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            var actual = richString.Substring(0);

            Assert.AreEqual(richString.First(), actual.First());

            Assert.AreEqual(1, actual.Count);

            actual.First().SetBold();

            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().First().Bold);
        }

        [Test]
        public void Substring_All_From_ThreeStrings()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(0);

            Assert.AreEqual(richString.ElementAt(0), actual.ElementAt(0));
            Assert.AreEqual(richString.ElementAt(1), actual.ElementAt(1));
            Assert.AreEqual(richString.ElementAt(2), actual.ElementAt(2));

            Assert.AreEqual(3, actual.Count);
            Assert.AreEqual(3, richString.Count);

            actual.First().SetBold();

            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().First().Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().Last().Bold);
        }

        [Test]
        public void Substring_From_OneString_End()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            var actual = richString.Substring(2);

            Assert.AreEqual(1, actual.Count); // substring was in one piece

            Assert.AreEqual(2, richString.Count); // The text was split because of the substring

            Assert.AreEqual("llo", actual.First().Text);

            Assert.AreEqual("He", richString.First().Text);
            Assert.AreEqual("llo", richString.Last().Text);

            actual.First().SetBold();

            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().First().Bold);
            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().Last().Bold);

            richString.Last().SetItalic();

            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().First().Italic);
            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().Last().Italic);

            Assert.AreEqual(true, actual.First().Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().First().FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().Last().FontSize);

            Assert.AreEqual(20, actual.First().FontSize);
        }

        [Test]
        public void Substring_From_OneString_Middle()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            var actual = richString.Substring(2, 2);

            Assert.AreEqual(1, actual.Count); // substring was in one piece

            Assert.AreEqual(3, richString.Count); // The text was split because of the substring

            Assert.AreEqual("ll", actual.First().Text);

            Assert.AreEqual("He", richString.First().Text);
            Assert.AreEqual("ll", richString.ElementAt(1).Text);
            Assert.AreEqual("o", richString.Last().Text);

            actual.First().SetBold();

            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().First().Bold);
            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().Last().Bold);

            richString.Last().SetItalic();

            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().First().Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Italic);
            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().Last().Italic);

            Assert.AreEqual(false, actual.First().Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().First().FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().Last().FontSize);

            Assert.AreEqual(20, actual.First().FontSize);
        }

        [Test]
        public void Substring_From_OneString_Start()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            var actual = richString.Substring(0, 2);

            Assert.AreEqual(1, actual.Count); // substring was in one piece

            Assert.AreEqual(2, richString.Count); // The text was split because of the substring

            Assert.AreEqual("He", actual.First().Text);

            Assert.AreEqual("He", richString.First().Text);
            Assert.AreEqual("llo", richString.Last().Text);

            actual.First().SetBold();

            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().First().Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().Last().Bold);

            richString.Last().SetItalic();

            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().First().Italic);
            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().Last().Italic);

            Assert.AreEqual(false, actual.First().Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().First().FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().Last().FontSize);

            Assert.AreEqual(20, actual.First().FontSize);
        }

        [Test]
        public void Substring_From_ThreeStrings_End1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(21);

            Assert.AreEqual(1, actual.Count); // substring was in one piece

            Assert.AreEqual(4, richString.Count); // The text was split because of the substring

            Assert.AreEqual("bors!", actual.First().Text);

            Assert.AreEqual("Good Morning", richString.ElementAt(0).Text);
            Assert.AreEqual(" my ", richString.ElementAt(1).Text);
            Assert.AreEqual("neigh", richString.ElementAt(2).Text);
            Assert.AreEqual("bors!", richString.ElementAt(3).Text);

            actual.First().SetBold();

            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(0).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Bold);
            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(3).Bold);

            richString.Last().SetItalic();

            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(0).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Italic);
            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(3).Italic);

            Assert.AreEqual(true, actual.First().Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize);

            Assert.AreEqual(20, actual.First().FontSize);
        }

        [Test]
        public void Substring_From_ThreeStrings_End2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(13);

            Assert.AreEqual(2, actual.Count);

            Assert.AreEqual(4, richString.Count); // The text was split because of the substring

            Assert.AreEqual("my ", actual.ElementAt(0).Text);
            Assert.AreEqual("neighbors!", actual.ElementAt(1).Text);

            Assert.AreEqual("Good Morning", richString.ElementAt(0).Text);
            Assert.AreEqual(" ", richString.ElementAt(1).Text);
            Assert.AreEqual("my ", richString.ElementAt(2).Text);
            Assert.AreEqual("neighbors!", richString.ElementAt(3).Text);

            actual.ElementAt(1).SetBold();

            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(0).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Bold);
            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(3).Bold);

            richString.Last().SetItalic();

            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(0).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Italic);
            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(3).Italic);

            Assert.AreEqual(false, actual.ElementAt(0).Italic);
            Assert.AreEqual(true, actual.ElementAt(1).Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize);

            Assert.AreEqual(20, actual.ElementAt(0).FontSize);
            Assert.AreEqual(20, actual.ElementAt(1).FontSize);
        }

        [Test]
        public void Substring_From_ThreeStrings_Mid1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(5, 10);

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
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(5, 15);

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
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(0, 4);

            Assert.AreEqual(1, actual.Count); // substring was in one piece

            Assert.AreEqual(4, richString.Count); // The text was split because of the substring

            Assert.AreEqual("Good", actual.First().Text);

            Assert.AreEqual("Good", richString.ElementAt(0).Text);
            Assert.AreEqual(" Morning", richString.ElementAt(1).Text);
            Assert.AreEqual(" my ", richString.ElementAt(2).Text);
            Assert.AreEqual("neighbors!", richString.ElementAt(3).Text);

            actual.First().SetBold();

            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(0).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(3).Bold);

            richString.First().SetItalic();

            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(0).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(3).Italic);

            Assert.AreEqual(true, actual.First().Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize);

            Assert.AreEqual(20, actual.First().FontSize);
        }

        [Test]
        public void Substring_From_ThreeStrings_Start2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(0, 15);

            Assert.AreEqual(2, actual.Count);

            Assert.AreEqual(4, richString.Count); // The text was split because of the substring

            Assert.AreEqual("Good Morning", actual.ElementAt(0).Text);
            Assert.AreEqual(" my", actual.ElementAt(1).Text);

            Assert.AreEqual("Good Morning", richString.ElementAt(0).Text);
            Assert.AreEqual(" my", richString.ElementAt(1).Text);
            Assert.AreEqual(" ", richString.ElementAt(2).Text);
            Assert.AreEqual("neighbors!", richString.ElementAt(3).Text);

            actual.ElementAt(1).SetBold();

            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(0).Bold);
            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Bold);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(3).Bold);

            richString.First().SetItalic();

            Assert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(0).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Italic);
            Assert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(3).Italic);

            Assert.AreEqual(true, actual.ElementAt(0).Italic);
            Assert.AreEqual(false, actual.ElementAt(1).Italic);

            richString.SetFontSize(20);

            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize);
            Assert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize);

            Assert.AreEqual(20, actual.ElementAt(0).FontSize);
            Assert.AreEqual(20, actual.ElementAt(1).FontSize);
        }

        [Test]
        public void Substring_IndexOutsideRange1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            Assert.That(() => richString.Substring(50), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText("World");

            Assert.That(() => richString.Substring(50), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange3()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            Assert.That(() => richString.Substring(1, 10), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange4()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

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
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World");
            var expected = "Hello World";
            var actual = richString.ToString();
            Assert.AreEqual(expected, actual);

            richString.AddText("!");
            expected = "Hello World!";
            actual = richString.ToString();
            Assert.AreEqual(expected, actual);

            richString.ClearText();
            expected = string.Empty;
            actual = richString.ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test(Description = "See #1361")]
        public void CanClearInlinedRichText()
        {
            using var outputStream = new MemoryStream();
            using (var inputStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\InlinedRichText\ChangeRichText\inputfile.xlsx")))
            using (var workbook = new XLWorkbook(inputStream))
            {
                workbook.Worksheets.First().Cell("A1").Value = "";
                workbook.SaveAs(outputStream);
            }

            using var wb = new XLWorkbook(outputStream);
            Assert.AreEqual("", wb.Worksheets.First().Cell("A1").Value);
        }

        [Test]
        public void CanChangeInlinedRichText()
        {
            static void testRichText(IXLRichText richText)
            {
                Assert.IsNotNull(richText);
                Assert.IsTrue(richText.Any());
                Assert.AreEqual("3", richText.ElementAt(2).Text);
                Assert.AreEqual(XLColor.Red, richText.ElementAt(2).FontColor);
            }

            using var outputStream = new MemoryStream();
            using (var inputStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\InlinedRichText\ChangeRichText\inputfile.xlsx")))
            using (var workbook = new XLWorkbook(inputStream))
            {
                var richText = workbook.Worksheets.First().Cell("A1").GetRichText();
                testRichText(richText);
                richText.AddText(" - changed");
                workbook.SaveAs(outputStream);
            }

            using var wb = new XLWorkbook(outputStream);
            var cell = wb.Worksheets.First().Cell("A1");
            Assert.IsFalse(cell.ShareString);
            Assert.IsTrue(cell.HasRichText);
            var rt = cell.GetRichText();
            Assert.AreEqual("Year (range: 3 yrs) - changed", rt.ToString());
            testRichText(rt);
        }

        [Test]
        public void ClearInlineRichTextWhenRelevant()
        {
            var expectedFilePath = @"Other\InlinedRichText\ChangeRichTextToFormula\output.xlsx";

            using var ms = new MemoryStream();
            TestHelper.CreateAndCompare(() =>
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet();
                    var cell = ws.FirstCell();

                    cell.GetRichText().AddText("Bold").SetBold().AddText(" and red").SetBold().SetFontColor(XLColor.Red);
                    cell.ShareString = false;

                    //wb.SaveAs(ms);
                    wb.SaveAs(ms);
                }
                ms.Seek(0, SeekOrigin.Begin);

                var wb2 = new XLWorkbook(ms);
                {
                    var ws = wb2.Worksheets.First();
                    var cell = ws.FirstCell();

                    cell.FormulaA1 = "=1 + 2";
                    wb2.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("../../../", "Resource", expectedFilePath));
                //File.WriteAllBytes(expectedFileInVsSolution, ms.ToArray());

                return wb2;
            }, expectedFilePath);
        }
    }
}