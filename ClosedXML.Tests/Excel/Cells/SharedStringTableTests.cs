using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Cells
{
    [TestFixture]
    public class SharedStringTableTests
    {
        [Test]
        public void CanAccessTextThroughId()
        {
            var sst = new SharedStringTable();
            var id = sst.IncreaseRef("test");
            Assert.AreEqual("test", sst[id]);
            Assert.AreEqual(1, sst.Count);
        }

        [Test]
        public void TextsWithoutReferenceAreRemoved()
        {
            var sst = new SharedStringTable();
            var id = sst.IncreaseRef("test");
            sst.DecreaseRef(id);

            Assert.AreEqual(0, sst.Count);
            Assert.That(() => _ = sst[id], Throws.ArgumentException.With.Message.EqualTo("Id 0 has no text."));
        }

        [Test]
        public void TextReferencedByMultipleThingsIsNotFreedUntilAllAreRelease()
        {
            const string text = "test";
            var sst = new SharedStringTable();
            var id = sst.IncreaseRef(text);

            sst.IncreaseRef(text);
            Assert.AreEqual(text, sst[id]);
            Assert.AreEqual(1, sst.Count);

            sst.DecreaseRef(id);
            Assert.AreEqual(text, sst[id]);
            Assert.AreEqual(1, sst.Count);

            sst.IncreaseRef(text);
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
            sst.IncreaseRef("zero");
            var originalId = sst.IncreaseRef("original");
            var laterId = sst.IncreaseRef("two");

            Assert.That(laterId, Is.GreaterThan(originalId));

            sst.DecreaseRef(originalId);
            Assert.Throws<ArgumentException>(() => _ = sst[originalId]);

            var replacementId = sst.IncreaseRef("replacement");
            Assert.AreEqual(originalId, replacementId);
            Assert.AreEqual("replacement", sst[replacementId]);
        }

        [Test]
        public void DereferencingFreedIdThrows()
        {
            var sst = new SharedStringTable();
            var id = sst.IncreaseRef("test");
            sst.DecreaseRef(id);
            Assert.Throws<InvalidOperationException>(() => sst.DecreaseRef(id));
        }
    }
}
