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
            Assert.That(reader.GetPropertyName(0), Is.Null);
        }

        [Test]
        public void CanGetPropertiesCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.That(reader.GetPropertiesCount(), Is.EqualTo(3));
        }

        [Test]
        public void CanGetRecordsCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.That(reader.GetRecords().Count(), Is.EqualTo(2));
        }

        [Test]
        public void CanReadValues()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            var result = reader.GetRecords();

            Assert.Multiple(() =>
            {
                Assert.That(result.First().First(), Is.EqualTo(1));
                Assert.That(result.First().Last(), Is.EqualTo(3));
                Assert.That(result.Last().First(), Is.EqualTo(4));
                Assert.That(result.Last().Last(), Is.EqualTo(6));
            });
        }
    }
}
