using ClosedXML.Excel.InsertData;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ClosedXML.Tests.Excel.InsertData
{
    public class SimpleNullableTypeReaderTests
    {
        private readonly int?[] _data = { 1, 2, null };

        [TestCaseSource(nameof(SimpleNullableSourceNames))]
        public string CanGetPropertyName<T>(IEnumerable<T> data)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return reader.GetPropertyName(0);
        }

        private static IEnumerable<TestCaseData> SimpleNullableSourceNames
        {
            get
            {
                yield return new TestCaseData(new int?[] { 1, 2, null }).Returns("Int32");
                yield return new TestCaseData(new List<double?> { 1.0, 2.0, null }).Returns("Double");
                yield return new TestCaseData(new decimal?[] { 1.0m, 2.0m, null }).Returns("Decimal");
                yield return new TestCaseData(new char?[] { 'A', 'B', null }).Returns("Char");
                yield return new TestCaseData(new DateTime?[] { new DateTime(2020, 1, 1), null }).Returns("DateTime");
            }
        }

        [Test]
        public void CanGetPropertiesCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.That(reader.GetPropertiesCount(), Is.EqualTo(1));
        }

        [Test]
        public void CanGetRecordsCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            Assert.That(reader.GetRecords().Count(), Is.EqualTo(3));
        }

        [Test]
        public void CanReadValues()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
            var result = reader.GetRecords();

            Assert.Multiple(() =>
            {
                Assert.That(result.First().Single(), Is.EqualTo(1));
                Assert.That(result.Last().Single(), Is.EqualTo(Blank.Value));
            });
        }
    }
}
