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

            using var connection = new SqlConnection(_connectionString);
            using var command = new SqlCommand(queryString, connection);
            try
            {
                connection.Open();
            }
            catch
            {
                Assert.Ignore("Could not connect to localdb");
            }

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                yield return reader;
            }
        }

        [Test]
        public void CanGetPropertyName()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
            Assert.Multiple(() =>
            {
                Assert.That(reader.GetPropertyName(0), Is.EqualTo("StringValue"));
                Assert.That(reader.GetPropertyName(1), Is.EqualTo("NumericValue"));
            });
        }

        [Test]
        public void CanGetPropertiesCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
            Assert.That(reader.GetPropertiesCount(), Is.EqualTo(2));
        }

        [Test]
        public void CanGetRecordsCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
            Assert.That(reader.GetRecords().Count(), Is.EqualTo(3));
        }

        [Test]
        public void CanGetData()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
            var result = reader.GetRecords().ToArray();

            Assert.Multiple(() =>
            {
                Assert.That(result.First().First(), Is.EqualTo("Value 1"));
                Assert.That(result.First().Last(), Is.EqualTo(100));
                Assert.That(result.Last().First(), Is.EqualTo("Value 3"));
                Assert.That(result.Last().Last(), Is.EqualTo(300));
            });
        }
    }
}
