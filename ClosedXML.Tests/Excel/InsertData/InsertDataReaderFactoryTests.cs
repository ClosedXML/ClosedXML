using ClosedXML.Excel.InsertData;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ClosedXML.Tests.Excel.InsertData
{
    public class InsertDataReaderFactoryTests
    {
        [Test]
        public void CanInstantiateFactory()
        {
            var factory = InsertDataReaderFactory.Instance;

            Assert.IsNotNull(factory);
            Assert.AreSame(factory, InsertDataReaderFactory.Instance);
        }

        [TestCaseSource(nameof(SimpleSources))]
        public void CanCreateSimpleReader(IEnumerable data)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);

            Assert.IsInstanceOf<SimpleTypeReader>(reader);
        }

        private static IEnumerable<object> SimpleSources
        {
            get
            {
                yield return new[] { 1, 2, 3 };
                yield return new List<double> { 1.0, 2.0, 3.0 };
                yield return new[] { "A", "B", "C" };
                yield return new[] { "A", "B", "C" }.Cast<object>();
                yield return new[] { 'A', 'B', 'C' };
            }
        }

        [TestCaseSource(nameof(SimpleNullableSources))]
        public void CanCreateSimpleNullableReader(IEnumerable data)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);

            Assert.IsInstanceOf<SimpleNullableTypeReader>(reader);
        }

        private static IEnumerable<object> SimpleNullableSources
        {
            get
            {
                yield return new int?[] { 1, 2, null };
                yield return new List<double?> { 1.0, 2.0, null };
                yield return new char?[] { 'A', 'B', null };
                yield return new DateTime?[] { DateTime.MinValue, DateTime.MaxValue, null };
            }
        }

        [TestCaseSource(nameof(ArraySources))]
        public void CanCreateArrayReader<T>(IEnumerable<T> data)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);

            Assert.IsInstanceOf<ArrayReader>(reader);
        }

        private static IEnumerable<object> ArraySources
        {
            get
            {
                yield return new int[][]
                {
                    new[] {1, 2, 3},
                    new[] {4, 5, 6}
                };
                yield return new List<List<double>> { new List<double> { 1.0, 2.0, 3.0 } };
                yield return (new int[][]
                {
                    new[] {1, 2, 3},
                    new[] {4, 5, 6}
                }).AsEnumerable();
                yield return new Array[]
                {
                    Array.CreateInstance(typeof(decimal), 5),
                    Array.CreateInstance(typeof(decimal), 5),
                };
            }
        }

        [Test]
        public void CanCreateArrayReaderFromIEnumerableOfIEnumerables()
        {
            IEnumerable<IEnumerable> data = new List<IEnumerable>
            {
                new[] {1, 2, 3}.AsEnumerable(),
                new[] {1.0, 2.0, 3.0}.AsEnumerable(),
            };
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);

            Assert.IsInstanceOf<ArrayReader>(reader);
        }

        [Test]
        public void CanCreateSimpleReaderFromIEnumerableOfString()
        {
            IEnumerable<string> data = new[]
            {
                "String 1",
                "String 2",
            };
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);

            Assert.IsInstanceOf<SimpleTypeReader>(reader);
        }

        [Test]
        public void CanCreateDataTableReader()
        {
            var dt = new DataTable();
            var reader = InsertDataReaderFactory.Instance.CreateReader(dt);

            Assert.IsInstanceOf<ClosedXML.Excel.InsertData.DataTableReader>(reader);
        }

        [Test]
        public void CanCreateDataRecordReader()
        {
            var dataRecords = new IDataRecord[0];
            var reader = InsertDataReaderFactory.Instance.CreateReader(dataRecords);
            Assert.IsInstanceOf<DataRecordReader>(reader);
        }

        [Test]
        public void CanCreateObjectReader()
        {
            var entities = new TestEntity[0];
            var reader = InsertDataReaderFactory.Instance.CreateReader(entities);
            Assert.IsInstanceOf<ObjectReader>(reader);
        }

        [Test]
        public void CanCreateObjectReaderForStruct()
        {
            var entities = new TestStruct[0];
            var reader = InsertDataReaderFactory.Instance.CreateReader(entities);
            Assert.IsInstanceOf<ObjectReader>(reader);
        }

        [Test]
        public void CanCreateUntypedObjectReader()
        {
            var entities = new ArrayList(new object[]
            {
                new TestEntity(),
                "123",
            });
            var reader = InsertDataReaderFactory.Instance.CreateReader(entities);
            Assert.IsInstanceOf<UntypedObjectReader>(reader);
        }

        private class TestEntity { }

        private struct TestStruct { }
    }
}
