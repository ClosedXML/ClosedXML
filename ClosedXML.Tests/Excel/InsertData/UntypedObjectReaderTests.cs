using ClosedXML.Excel.InsertData;
using NUnit.Framework;
using System.Collections;
using System.Linq;

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
            Assert.AreEqual(expectedPropertyName, actualPropertyName);
        }

        [Test]
        public void CanGetPropertiesCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.AreEqual(4, reader.GetPropertiesCount());
        }

        [Test]
        public void CanGetRecordsCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.AreEqual(9, reader.GetRecordsCount());
        }

        [Test]
        public void CanGetData()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);

            var result = reader.GetData().ToArray();

            Assert.AreEqual(new object[] { null }, result[0]);
            Assert.AreEqual(new object[] { "Value 2", "Value 1", 4, 3 }, result[1]);
            Assert.AreEqual(new object[] { null }, result[2]);
            Assert.AreEqual(new object[] { null }, result[3]);
            Assert.AreEqual(new object[] { null }, result[4]);
            Assert.AreEqual(new object[] { 1, 2, 3 }, result[5]);
            Assert.AreEqual(new object[] { 4, 5, 6, 7 }, result[6]);
            Assert.AreEqual(new object[] { "Separator" }, result[7]);
            Assert.AreEqual(new object[] { "Value 9", "Value 10" }, result[8]);
        }
    }
}
