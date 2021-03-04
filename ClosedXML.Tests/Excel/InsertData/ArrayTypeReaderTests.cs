using ClosedXML.Excel.InsertData;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.InsertData
{
    public class ArrayTypeReaderTests
    {
        private readonly int[][] _data = new int[][]
        {
            new[] {1, 2, 3},
            new[] {4, 5, 6}
        };

        [Test]
        public void GetPropertyNameReturnsNull()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.IsNull(reader.GetPropertyName(0));
        }

        [Test]
        public void CanGetPropertiesCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.AreEqual(3, reader.GetPropertiesCount());
        }

        [Test]
        public void CanGetRecordsCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.AreEqual(2, reader.GetRecordsCount());
        }

        [Test]
        public void CanReadValues()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            var result = reader.GetData();

            Assert.AreEqual(1, result.First().First());
            Assert.AreEqual(3, result.First().Last());
            Assert.AreEqual(4, result.Last().First());
            Assert.AreEqual(6, result.Last().Last());
        }
    }
}
