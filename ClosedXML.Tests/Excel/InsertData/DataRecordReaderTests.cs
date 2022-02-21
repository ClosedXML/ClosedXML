using ClosedXML.Excel.InsertData;
using NUnit.Framework;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace ClosedXML.Tests.Excel.InsertData
{
    public class DataRecordReaderTests
    {
        private readonly string _connectionString = @"Data Source=(localdb)\MSSQLLocalDB;Integrated Security=True;Connect Timeout=1";

        private IEnumerable<IDataRecord> GetData()
        {
            const string queryString = @"
            select 'Value 1' as StringValue, 100 as NumericValue
            union all
            select 'Value 2', 200
            union all
            select 'Value 3', 300";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(queryString, connection))
            {
                try
                {
                    connection.Open();
                }
                catch
                {
                    Assert.Ignore("Could not connect to localdb");
                }

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        yield return reader;
                    }
                }
            }
        }

        [Test]
        public void CanGetPropertyName()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
            Assert.AreEqual("StringValue", reader.GetPropertyName(0));
            Assert.AreEqual("NumericValue", reader.GetPropertyName(1));
        }

        [Test]
        public void CanGetPropertiesCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
            Assert.AreEqual(2, reader.GetPropertiesCount());
        }

        [Test]
        public void CanGetRecordsCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
            Assert.AreEqual(3, reader.GetRecordsCount());
        }

        [Test]
        public void CanGetData()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
            var result = reader.GetData().ToArray();

            Assert.AreEqual("Value 1", result.First().First());
            Assert.AreEqual(100, result.First().Last());
            Assert.AreEqual("Value 3", result.Last().First());
            Assert.AreEqual(300, result.Last().Last());
        }
    }
}
