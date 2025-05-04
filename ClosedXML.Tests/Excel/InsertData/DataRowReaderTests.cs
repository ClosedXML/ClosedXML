using ClosedXML.Excel.InsertData;
using NUnit.Framework;
using System.Data;
using System.Linq;

namespace ClosedXML.Tests.Excel.InsertData
{
    public class DataRowReaderTests
    {
        private readonly DataTable _data;

        public DataRowReaderTests()
        {
            _data = new DataTable();
            _data.Columns.Add("Last name");
            _data.Columns.Add("First name");
            _data.Columns.Add("Age", typeof(int));

            _data.Rows.Add("Smith", "John", 33);
            _data.Rows.Add("Ivanova", "Olga", 25);
        }

        [Test]
        public void CanGetPropertyName()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.Multiple(() =>
            {
                Assert.That(reader.GetPropertyName(0), Is.EqualTo("Last name"));
                Assert.That(reader.GetPropertyName(1), Is.EqualTo("First name"));
                Assert.That(reader.GetPropertyName(2), Is.EqualTo("Age"));
            });
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
        public void CanReadValue()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            var result = reader.GetRecords();

            Assert.Multiple(() =>
            {
                Assert.That(result.First().First(), Is.EqualTo("Smith"));
                Assert.That(result.First().Last(), Is.EqualTo(33));
                Assert.That(result.Last().First(), Is.EqualTo("Ivanova"));
                Assert.That(result.Last().Last(), Is.EqualTo(25));
            });
        }
        
        [OneTimeTearDown]
        public void Cleanup()
        {
            _data.Dispose();
        }
    }
}
