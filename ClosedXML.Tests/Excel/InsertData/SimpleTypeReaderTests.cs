using ClosedXML.Excel.InsertData;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests.Excel.InsertData
{
    public class SimpleTypeReaderTests
    {
        private readonly int[] _data = { 1, 2, 3 };

        [TestCaseSource(nameof(SimpleSourceNames))]
        public string CanGetPropertyName<T>(IEnumerable<T> data)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return reader.GetPropertyName(0);
        }

        private static IEnumerable<TestCaseData> SimpleSourceNames
        {
            get
            {
                yield return new TestCaseData(new[] { 1, 2, 3 }).Returns("Int32");
                yield return new TestCaseData(new List<double> { 1.0, 2.0, 3.0 }).Returns("Double");
                yield return new TestCaseData(new[] { 1.0m, 2.0m, 3.0m }).Returns("Decimal");
                yield return new TestCaseData(arg: new[] { "A", "B", "C" }).Returns("String");
                yield return new TestCaseData(new[] { 'A', 'B', 'C' }).Returns("Char");
                yield return new TestCaseData(new[] { new DateTime(2020, 1, 1) }).Returns("DateTime");
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
                Assert.That(result.Last().Single(), Is.EqualTo(3));
            });
        }
    }
}
