using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests
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
            cell.CreateRichText().AddText("12");

            IXLRichText richText = cell.GetRichText();

            Assert.That(richText.ToString(), Is.EqualTo("12"));

            richText.AddText("34");

            Assert.That(cell.GetText(), Is.EqualTo("1234"));
        }

        /// <summary>
        ///     A test for AddText
        /// </summary>
        [Test]
        public void AddTextTest1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            IXLRichText richString = cell.CreateRichText();

            string text = "Hello";
            richString.AddText(text).SetBold().SetFontColor(XLColor.Red);

            Assert.Multiple(() =>
            {
                Assert.That(text, Is.EqualTo(cell.GetText()));
                Assert.That(true, Is.EqualTo(cell.GetRichText().First().Bold));
                Assert.That(XLColor.Red, Is.EqualTo(cell.GetRichText().First().FontColor));

                Assert.That(richString, Has.Count.EqualTo(1));
            });

            richString.AddText("World");
            Assert.That(text, Is.EqualTo(richString.First().Text), "Item in collection is not the same as the one returned");
        }

        [Test]
        public void AddTextTest2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            int number = 123;

            cell.SetValue(number).Style
                .Font.SetBold()
                .Font.SetFontColor(XLColor.Red);

            string text = number.ToString();

            Assert.Multiple(() =>
            {
                Assert.That(text, Is.EqualTo(cell.GetRichText().ToString()));
                Assert.That(true, Is.EqualTo(cell.GetRichText().First().Bold));
                Assert.That(XLColor.Red, Is.EqualTo(cell.GetRichText().First().FontColor));

                Assert.That(cell.GetRichText(), Has.Count.EqualTo(1));
            });

            cell.GetRichText().AddText("World");
            Assert.That(text, Is.EqualTo(cell.GetRichText().First().Text), "Item in collection is not the same as the one returned");
        }

        [Test]
        public void AddTextTest3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            int number = 123;
            cell.Value = number;
            cell.Style
                .Font.SetBold()
                .Font.SetFontColor(XLColor.Red);

            string text = number.ToString();

            Assert.Multiple(() =>
            {
                Assert.That(text, Is.EqualTo(cell.GetRichText().ToString()));
                Assert.That(true, Is.EqualTo(cell.GetRichText().First().Bold));
                Assert.That(XLColor.Red, Is.EqualTo(cell.GetRichText().First().FontColor));

                Assert.That(cell.GetRichText(), Has.Count.EqualTo(1));
            });

            cell.GetRichText().AddText("World");
            Assert.That(text, Is.EqualTo(cell.GetRichText().First().Text), "Item in collection is not the same as the one returned");
        }

        /// <summary>
        ///     A test for Clear
        /// </summary>
        [Test]
        public void ClearTest()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World!");

            richString.ClearText();
            string expected = string.Empty;
            string actual = richString.ToString();
            Assert.Multiple(() =>
            {
                Assert.That(actual, Is.EqualTo(expected));

                Assert.That(richString.Count, Is.EqualTo(0));
            });
        }

        [Test]
        public void CountTest()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World!");

            Assert.That(richString, Has.Count.EqualTo(3));
        }

        [Test]
        public void HasRichTextTest1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.GetRichText().AddText("123");

            Assert.That(cell.HasRichText, Is.True);

            cell.Value = "123";

            Assert.That(cell.HasRichText, Is.False);

            cell.GetRichText().AddText("123");

            Assert.That(cell.HasRichText, Is.True);

            cell.Value = 123;

            Assert.That(cell.HasRichText, Is.False);

            cell.GetRichText().AddText("123");

            Assert.That(cell.HasRichText, Is.True);

            cell.SetValue("123");

            Assert.That(cell.HasRichText, Is.False);
        }

        /// <summary>
        ///     A test for Characters
        /// </summary>
        [Test]
        public void Substring_All_From_OneString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            IXLFormattedText<IXLRichText> actual = richString.Substring(0);

            Assert.Multiple(() =>
            {
                Assert.That(actual.First(), Is.EqualTo(richString.First()));

                Assert.That(actual, Has.Count.EqualTo(1));
            });

            actual.First().SetBold();

            Assert.That(ws.Cell(1, 1).GetRichText().First().Bold, Is.True);
        }

        [Test]
        public void Substring_All_From_ThreeStrings()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(0);

            Assert.Multiple(() =>
            {
                Assert.That(actual.ElementAt(0), Is.EqualTo(richString.ElementAt(0)));
                Assert.That(actual.ElementAt(1), Is.EqualTo(richString.ElementAt(1)));
                Assert.That(actual.ElementAt(2), Is.EqualTo(richString.ElementAt(2)));

                Assert.That(actual, Has.Count.EqualTo(3));
                Assert.That(richString, Has.Count.EqualTo(3));
            });

            actual.First().SetBold();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().First().Bold, Is.True);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold, Is.False);
            });
        }

        [Test]
        public void Substring_From_OneString_End()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            IXLFormattedText<IXLRichText> actual = richString.Substring(2);

            Assert.Multiple(() =>
            {
                Assert.That(actual, Has.Count.EqualTo(1)); // substring was in one piece

                Assert.That(richString, Has.Count.EqualTo(2)); // The text was split because of the substring

                Assert.That(actual.First().Text, Is.EqualTo("llo"));
            });

            Assert.Multiple(() =>
            {
                Assert.That(richString.First().Text, Is.EqualTo("He"));
                Assert.That(richString.Last().Text, Is.EqualTo("llo"));
            });

            actual.First().SetBold();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().First().Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold, Is.True);
            });

            richString.Last().SetItalic();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().First().Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().Last().Italic, Is.True);

                Assert.That(actual.First().Italic, Is.True);
            });

            richString.SetFontSize(20);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().First().FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().Last().FontSize, Is.EqualTo(20));

                Assert.That(actual.First().FontSize, Is.EqualTo(20));
            });
        }

        [Test]
        public void Substring_From_OneString_Middle()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            IXLFormattedText<IXLRichText> actual = richString.Substring(2, 2);

            Assert.Multiple(() =>
            {
                Assert.That(actual, Has.Count.EqualTo(1)); // substring was in one piece

                Assert.That(richString, Has.Count.EqualTo(3)); // The text was split because of the substring

                Assert.That(actual.First().Text, Is.EqualTo("ll"));
            });

            Assert.Multiple(() =>
            {
                Assert.That(richString.First().Text, Is.EqualTo("He"));
                Assert.That(richString.ElementAt(1).Text, Is.EqualTo("ll"));
                Assert.That(richString.Last().Text, Is.EqualTo("o"));
            });

            actual.First().SetBold();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().First().Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.True);
                Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold, Is.False);
            });

            richString.Last().SetItalic();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().First().Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().Last().Italic, Is.True);

                Assert.That(actual.First().Italic, Is.False);
            });

            richString.SetFontSize(20);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().First().FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().Last().FontSize, Is.EqualTo(20));

                Assert.That(actual.First().FontSize, Is.EqualTo(20));
            });
        }

        [Test]
        public void Substring_From_OneString_Start()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            IXLFormattedText<IXLRichText> actual = richString.Substring(0, 2);

            Assert.Multiple(() =>
            {
                Assert.That(actual, Has.Count.EqualTo(1)); // substring was in one piece

                Assert.That(richString, Has.Count.EqualTo(2)); // The text was split because of the substring

                Assert.That(actual.First().Text, Is.EqualTo("He"));
            });

            Assert.Multiple(() =>
            {
                Assert.That(richString.First().Text, Is.EqualTo("He"));
                Assert.That(richString.Last().Text, Is.EqualTo("llo"));
            });

            actual.First().SetBold();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().First().Bold, Is.True);
                Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold, Is.False);
            });

            richString.Last().SetItalic();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().First().Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().Last().Italic, Is.True);

                Assert.That(actual.First().Italic, Is.False);
            });

            richString.SetFontSize(20);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().First().FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().Last().FontSize, Is.EqualTo(20));

                Assert.That(actual.First().FontSize, Is.EqualTo(20));
            });
        }

        [Test]
        public void Substring_From_ThreeStrings_End1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(21);

            Assert.Multiple(() =>
            {
                Assert.That(actual, Has.Count.EqualTo(1)); // substring was in one piece

                Assert.That(richString, Has.Count.EqualTo(4)); // The text was split because of the substring

                Assert.That(actual.First().Text, Is.EqualTo("bors!"));
            });

            Assert.Multiple(() =>
            {
                Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good Morning"));
                Assert.That(richString.ElementAt(1).Text, Is.EqualTo(" my "));
                Assert.That(richString.ElementAt(2).Text, Is.EqualTo("neigh"));
                Assert.That(richString.ElementAt(3).Text, Is.EqualTo("bors!"));
            });

            actual.First().SetBold();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold, Is.True);
            });

            richString.Last().SetItalic();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic, Is.True);

                Assert.That(actual.First().Italic, Is.True);
            });

            richString.SetFontSize(20);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize, Is.EqualTo(20));

                Assert.That(actual.First().FontSize, Is.EqualTo(20));
            });
        }

        [Test]
        public void Substring_From_ThreeStrings_End2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(13);

            Assert.Multiple(() =>
            {
                Assert.That(actual, Has.Count.EqualTo(2));

                Assert.That(richString, Has.Count.EqualTo(4)); // The text was split because of the substring

                Assert.That(actual.ElementAt(0).Text, Is.EqualTo("my "));
                Assert.That(actual.ElementAt(1).Text, Is.EqualTo("neighbors!"));
            });

            Assert.Multiple(() =>
            {
                Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good Morning"));
                Assert.That(richString.ElementAt(1).Text, Is.EqualTo(" "));
                Assert.That(richString.ElementAt(2).Text, Is.EqualTo("my "));
                Assert.That(richString.ElementAt(3).Text, Is.EqualTo("neighbors!"));
            });

            actual.ElementAt(1).SetBold();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold, Is.True);
            });

            richString.Last().SetItalic();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic, Is.True);

                Assert.That(actual.ElementAt(0).Italic, Is.False);
                Assert.That(actual.ElementAt(1).Italic, Is.True);
            });

            richString.SetFontSize(20);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize, Is.EqualTo(20));

                Assert.That(actual.ElementAt(0).FontSize, Is.EqualTo(20));
                Assert.That(actual.ElementAt(1).FontSize, Is.EqualTo(20));
            });
        }

        [Test]
        public void Substring_From_ThreeStrings_Mid1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(5, 10);

            Assert.Multiple(() =>
            {
                Assert.That(actual, Has.Count.EqualTo(2));

                Assert.That(richString, Has.Count.EqualTo(5)); // The text was split because of the substring

                Assert.That(actual.ElementAt(0).Text, Is.EqualTo("Morning"));
                Assert.That(actual.ElementAt(1).Text, Is.EqualTo(" my"));
            });

            Assert.Multiple(() =>
            {
                Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good "));
                Assert.That(richString.ElementAt(1).Text, Is.EqualTo("Morning"));
                Assert.That(richString.ElementAt(2).Text, Is.EqualTo(" my"));
                Assert.That(richString.ElementAt(3).Text, Is.EqualTo(" "));
                Assert.That(richString.ElementAt(4).Text, Is.EqualTo("neighbors!"));
            });
        }

        [Test]
        public void Substring_From_ThreeStrings_Mid2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(5, 15);

            Assert.Multiple(() =>
            {
                Assert.That(actual, Has.Count.EqualTo(3));

                Assert.That(richString, Has.Count.EqualTo(5)); // The text was split because of the substring

                Assert.That(actual.ElementAt(0).Text, Is.EqualTo("Morning"));
                Assert.That(actual.ElementAt(1).Text, Is.EqualTo(" my "));
                Assert.That(actual.ElementAt(2).Text, Is.EqualTo("neig"));
            });

            Assert.Multiple(() =>
            {
                Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good "));
                Assert.That(richString.ElementAt(1).Text, Is.EqualTo("Morning"));
                Assert.That(richString.ElementAt(2).Text, Is.EqualTo(" my "));
                Assert.That(richString.ElementAt(3).Text, Is.EqualTo("neig"));
                Assert.That(richString.ElementAt(4).Text, Is.EqualTo("hbors!"));
            });
        }

        [Test]
        public void Substring_From_ThreeStrings_Start1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(0, 4);

            Assert.Multiple(() =>
            {
                Assert.That(actual, Has.Count.EqualTo(1)); // substring was in one piece

                Assert.That(richString, Has.Count.EqualTo(4)); // The text was split because of the substring

                Assert.That(actual.First().Text, Is.EqualTo("Good"));
            });

            Assert.Multiple(() =>
            {
                Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good"));
                Assert.That(richString.ElementAt(1).Text, Is.EqualTo(" Morning"));
                Assert.That(richString.ElementAt(2).Text, Is.EqualTo(" my "));
                Assert.That(richString.ElementAt(3).Text, Is.EqualTo("neighbors!"));
            });

            actual.First().SetBold();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold, Is.True);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold, Is.False);
            });

            richString.First().SetItalic();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic, Is.True);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic, Is.False);

                Assert.That(actual.First().Italic, Is.True);
            });

            richString.SetFontSize(20);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize, Is.EqualTo(20));

                Assert.That(actual.First().FontSize, Is.EqualTo(20));
            });
        }

        [Test]
        public void Substring_From_ThreeStrings_Start2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            IXLFormattedText<IXLRichText> actual = richString.Substring(0, 15);

            Assert.Multiple(() =>
            {
                Assert.That(actual, Has.Count.EqualTo(2));

                Assert.That(richString, Has.Count.EqualTo(4)); // The text was split because of the substring

                Assert.That(actual.ElementAt(0).Text, Is.EqualTo("Good Morning"));
                Assert.That(actual.ElementAt(1).Text, Is.EqualTo(" my"));
            });

            Assert.Multiple(() =>
            {
                Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good Morning"));
                Assert.That(richString.ElementAt(1).Text, Is.EqualTo(" my"));
                Assert.That(richString.ElementAt(2).Text, Is.EqualTo(" "));
                Assert.That(richString.ElementAt(3).Text, Is.EqualTo("neighbors!"));
            });

            actual.ElementAt(1).SetBold();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.True);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold, Is.False);
            });

            richString.First().SetItalic();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic, Is.True);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic, Is.False);
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic, Is.False);

                Assert.That(actual.ElementAt(0).Italic, Is.True);
                Assert.That(actual.ElementAt(1).Italic, Is.False);
            });

            richString.SetFontSize(20);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize, Is.EqualTo(20));
                Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize, Is.EqualTo(20));

                Assert.That(actual.ElementAt(0).FontSize, Is.EqualTo(20));
                Assert.That(actual.ElementAt(1).FontSize, Is.EqualTo(20));
            });
        }

        [Test]
        public void Substring_IndexOutsideRange1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            Assert.That(() => richString.Substring(50), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText("World");

            Assert.That(() => richString.Substring(50), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            Assert.That(() => richString.Substring(1, 10), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange4()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText("World");

            Assert.That(() => richString.Substring(5, 20), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void CopyFrom_DoesCopy()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var original = ws.Cell(1, 1).GetRichText();
            original
                .AddText("Hello").SetFontSize(15).SetFontColor(XLColor.Red)
                .AddText("World").SetFontSize(7).SetFontColor(XLColor.Blue);

            var otherCell = ws.Cell(1, 2);
            var otherRichText = otherCell.GetRichText();
            otherRichText.CopyFrom(original);

            Assert.Multiple(() =>
            {
                Assert.That(otherCell.Value, Is.EqualTo("HelloWorld"));
                Assert.That(otherRichText, Has.Count.EqualTo(2));
            });
            Assert.Multiple(() =>
            {
                Assert.That(otherRichText.First().FontColor, Is.EqualTo(XLColor.Red));
                Assert.That(otherRichText.Last().FontColor, Is.EqualTo(XLColor.Blue));
            });
        }

        /// <summary>
        ///     A test for ToString
        /// </summary>
        [Test]
        public void ToStringTest()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World");
            string expected = "Hello World";
            string actual = richString.ToString();
            Assert.That(actual, Is.EqualTo(expected));

            richString.AddText("!");
            expected = "Hello World!";
            actual = richString.ToString();
            Assert.That(actual, Is.EqualTo(expected));

            richString.ClearText();
            expected = string.Empty;
            actual = richString.ToString();
            Assert.That(actual, Is.EqualTo(expected));
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

            using (var wb = new XLWorkbook(outputStream))
            {
                Assert.That(wb.Worksheets.First().Cell("A1").Value, Is.EqualTo(""));
            }
        }

        [Test]
        public void CanChangeInlinedRichText()
        {
            static void AssertRichText(IXLRichText richText)
            {
                Assert.IsNotNull(richText);
                Assert.Multiple(() =>
                {
                    Assert.That(richText.Any(), Is.True);
                    Assert.That(richText.ElementAt(2).Text, Is.EqualTo("3"));
                    Assert.That(richText.ElementAt(2).FontColor, Is.EqualTo(XLColor.Red));
                });
            }

            using var outputStream = new MemoryStream();
            using (var inputStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\InlinedRichText\ChangeRichText\inputfile.xlsx")))
            using (var workbook = new XLWorkbook(inputStream))
            {
                var richText = workbook.Worksheets.First().Cell("A1").GetRichText();
                AssertRichText(richText);
                richText.AddText(" - changed");
                workbook.SaveAs(outputStream);
            }

            using (var wb = new XLWorkbook(outputStream))
            {
                var cell = wb.Worksheets.First().Cell("A1");
                Assert.Multiple(() =>
                {
                    Assert.That(cell.ShareString, Is.False);
                    Assert.That(cell.HasRichText, Is.True);
                });
                var rt = cell.GetRichText();
                Assert.That(rt.ToString(), Is.EqualTo("Year (range: 3 yrs) - changed"));
                AssertRichText(rt);
            }
        }

        [Test]
        public void ClearInlineRichTextWhenRelevant()
        {
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

                return wb2;
            }, @"Other\InlinedRichText\ChangeRichTextToFormula\output.xlsx");
        }

        [Test]
        public void RichTextChangesContentOfItsCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var cell = ws.Cell(1, 1);
            var richText = cell.GetRichText();

            Assert.That(richText.Text, Is.EqualTo(cell.Value));

            richText.AddText("Hello");
            Assert.That("Hello", Is.EqualTo(cell.Value));

            var world = richText.AddText(" World");
            Assert.That("Hello World", Is.EqualTo(cell.Value));

            world.Text = " World!";
            Assert.That("Hello World!", Is.EqualTo(cell.Value));
            Assert.That("Hello World!", Is.EqualTo(cell.GetRichText().Text));

            richText.ClearText();
            Assert.That(string.Empty, Is.EqualTo(cell.Value));
        }

        [Test]
        public void RemovedRichTextFromCellCantBeChanged()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var cell = ws.Cell(1, 1);
            var richText = cell.GetRichText();
            cell.Value = 4;

            Assert.Throws<InvalidOperationException>(() => richText.AddText("Hello"), "The rich text isn't a content of a cell.");
        }

        [Test]
        public void MaintainWhitespaces()
        {
            const string textWithSpaces = "  元  気  ";
            const string phoneticsWithSpace = "  げ  ん  ";
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var richTextCell = ws.Cell(1, 1);
                var richText = richTextCell.GetRichText();
                richText.AddText(textWithSpaces);
                richText.Phonetics.Add(phoneticsWithSpace, 2, 3);

                wb.SaveAs(ms);
            }

            ms.Position = 0;

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                var richText = ws.Cell(1, 1).GetRichText();
                Assert.Multiple(() =>
                {
                    Assert.That(richText.First().Text, Is.EqualTo(textWithSpaces));
                    Assert.That(richText.Phonetics.First().Text, Is.EqualTo(phoneticsWithSpace));
                });
            }
        }

        [Test]
        public void Preserve_end_of_line_in_xml()
        {
            // When text run in a rich text contains end of line (regardless if CR, LF or CRLF),
            // the written element must be marked with xml:space="preserve". Excel would process
            // text differently (trim ect, see XML spec) and that means there would be a data
            // loss (trimmed ends of line). Another problem would be phonetic runs. They use indexes
            // to the text run, but if text would be trimmed, they might suddenly have out-of-bounds
            // values and Excel would try to repair the workbook.
            // The source files contains a text run with end of line at the start and end. It also
            // contains phonetic run for the kanji in the text that would be out-of-bounds if space
            // attribute there. The input is from Excel, output is by ClosedXML. Output must contain
            // the space attribute.
            TestHelper.LoadSaveAndCompare(
                @"Other\RichText\kanji-with-new-line-input.xlsx",
                @"Other\RichText\kanji-with-new-line-output.xlsx");
        }
    }
}
