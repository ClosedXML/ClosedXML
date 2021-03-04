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
            Assert.AreEqual("Last name", reader.GetPropertyName(0));
            Assert.AreEqual("First name", reader.GetPropertyName(1));
            Assert.AreEqual("Age", reader.GetPropertyName(2));
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
        public void CanReadValue()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            var result = reader.GetData();

            Assert.AreEqual("Smith", result.First().First());
            Assert.AreEqual(33, result.First().Last());
            Assert.AreEqual("Ivanova", result.Last().First());
            Assert.AreEqual(25, result.Last().Last());
        }
    }
}
