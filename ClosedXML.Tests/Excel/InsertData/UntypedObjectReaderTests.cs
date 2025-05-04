using ClosedXML.Excel.InsertData;
using NUnit.Framework;
using System.Collections;
using System.Linq;
using ClosedXML.Excel;

namespace ClosedXML.Tests.Excel.InsertData
{
    public class UntypedObjectReaderTests
    {
        private readonly ArrayList _data = new ArrayList(new object[]
            {
                null,
                new TablesTests.TestObjectWithAttributes
                {
                    Column1 = "Value 1",
                    Column2 = "Value 2",
                    UnOrderedColumn = 3,
                    MyField = 4,
                },
                null,
                null,
                null,
                new int[]{ 1, 2, 3},
                new int[]{ 4, 5, 6, 7},
                "Separator",

                new TablesTests.TestObjectWithoutAttributes
                {
                    Column1 = "Value 9",
                    Column2 = "Value 10"
                },
            });

        [TestCase(0, "FirstColumn")]
        [TestCase(1, "SecondColumn")]
        [TestCase(2, "SomeFieldNotProperty")]
        [TestCase(3, "UnOrderedColumn")]
        public void CanGetPropertyName(int propertyIndex, string expectedPropertyName)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            var actualPropertyName = reader.GetPropertyName(propertyIndex);
            Assert.That(actualPropertyName, Is.EqualTo(expectedPropertyName));
        }

        [Test]
        public void CanGetPropertiesCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.That(reader.GetPropertiesCount(), Is.EqualTo(4));
        }

        [Test]
        public void CanGetRecordsCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.That(reader.GetRecords().Count(), Is.EqualTo(9));
        }

        [Test]
        public void CanGetData()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);

            var result = reader.GetRecords().ToArray();

            Assert.Multiple(() =>
            {
                Assert.That(result[0], Is.EqualTo(new XLCellValue[] { Blank.Value }));
                Assert.That(result[1], Is.EqualTo(new XLCellValue[] { "Value 2", "Value 1", 4, 3 }));
                Assert.That(result[2], Is.EqualTo(new XLCellValue[] { Blank.Value }));
                Assert.That(result[3], Is.EqualTo(new XLCellValue[] { Blank.Value }));
                Assert.That(result[4], Is.EqualTo(new XLCellValue[] { Blank.Value }));
                Assert.That(result[5], Is.EqualTo(new XLCellValue[] { 1, 2, 3 }));
                Assert.That(result[6], Is.EqualTo(new XLCellValue[] { 4, 5, 6, 7 }));
                Assert.That(result[7], Is.EqualTo(new XLCellValue[] { "Separator" }));
                Assert.That(result[8], Is.EqualTo(new XLCellValue[] { "Value 9", "Value 10" }));
            });
        }
    }
}
