using System;
using System.Text;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Cells
{
    [TestFixture]
    public class SharedStringTableTests
    {
        [Test]
        public void SameStringIsNotStoredTwice()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet();
            var ws2 = wb.AddWorksheet();
            var txt1 = "Hello";
            var txt2 = new StringBuilder("Hel").Append("lo").ToString();
            Assert.AreNotSame(txt1, txt2);

            ws1.Cell(1, 1).Value = txt1;
            ws2.Cell(1, 1).Value = txt2;

            Assert.AreSame(ws1.Cell(1, 1).Value.GetText(), ws2.Cell(1, 1).Value.GetText());
        }

        [Test]
        public void CanAccessTextThroughId()
        {
            var sst = new SharedStringTable();
            var id = sst.IncreaseRef("test", false);
            Assert.AreEqual("test", sst[id]);
            Assert.AreEqual(1, sst.Count);
        }

        [Test]
        public void TextsWithoutReferenceAreRemoved()
        {
            var sst = new SharedStringTable();
            var id = sst.IncreaseRef("test", false);
            sst.DecreaseRef(id);

            Assert.AreEqual(0, sst.Count);
            Assert.That(() => _ = sst[id], Throws.ArgumentException.With.Message.EqualTo("Id 0 has no text."));
        }

        [Test]
        public void TextReferencedByMultipleThingsIsNotFreedUntilAllAreRelease()
        {
            const string text = "test";
            var sst = new SharedStringTable();
            var id = sst.IncreaseRef(text, false);

            sst.IncreaseRef(text, false);
            Assert.AreEqual(text, sst[id]);
            Assert.AreEqual(1, sst.Count);

            sst.DecreaseRef(id);
            Assert.AreEqual(text, sst[id]);
            Assert.AreEqual(1, sst.Count);

            sst.IncreaseRef(text, false);
            Assert.AreEqual(text, sst[id]);
            Assert.AreEqual(1, sst.Count);

            sst.DecreaseRef(id);
            Assert.AreEqual(text, sst[id]);
            Assert.AreEqual(1, sst.Count);

            sst.DecreaseRef(id);
            Assert.Throws<ArgumentException>(() => _ = sst[id]);
        }

        [Test]
        public void FreedIdCanBeReusedForDifferentText()
        {
            var sst = new SharedStringTable();
            sst.IncreaseRef("zero", false);
            var originalId = sst.IncreaseRef("original", false);
            var laterId = sst.IncreaseRef("two", false);

            Assert.That(laterId, Is.GreaterThan(originalId));

            sst.DecreaseRef(originalId);
            Assert.Throws<ArgumentException>(() => _ = sst[originalId]);

            var replacementId = sst.IncreaseRef("replacement", false);
            Assert.AreEqual(originalId, replacementId);
            Assert.AreEqual("replacement", sst[replacementId]);
        }

        [Test]
        public void DereferencingFreedIdThrows()
        {
            var sst = new SharedStringTable();
            var id = sst.IncreaseRef("test", false);
            sst.DecreaseRef(id);
            Assert.Throws<InvalidOperationException>(() => sst.DecreaseRef(id));
        }

        [Test]
        public void StringItem_without_text_is_loaded_as_empty_text()
        {
            // PR#2218: A text cell that references self-closed <si/> tag in SST is loaded without
            // an error and is loaded as type TEXT. Although it's not very common, empty string is
            // a valid value of a cell.
            TestHelper.LoadAndAssert((_, ws) =>
            {
                // Check that type is a empty string, just like in Excel.
                Assert.AreEqual(2, ws.Evaluate("TYPE(B2)"));
                Assert.IsEmpty(ws.Cell("B2").GetText());
            }, @"Other\Cells\EmptySi.xlsx");
        }

        [Test]
        public void Empty_text_is_written_and_loaded_to_sst()
        {
            TestHelper.CreateSaveLoadAssert(
                (_, ws) =>
                {
                    ws.Cell("A1").Value = "Empty text cell (B1):";
                    ws.Cell("B1").Value = string.Empty;

                    ws.Cell("A2").Value = "Empty rich text";
                    ws.Cell("B2").CreateRichText().AddText(string.Empty);
                },
                (_, ws) =>
                {
                    Assert.AreEqual("", ws.Cell("B1").CachedValue);
                    Assert.AreEqual("", ws.Cell("B2").GetRichText().Text);
                },
                @"Other\Cells\EmptyText.xlsx");
        }
    }
}
