using ClosedXML.Excel.InsertData;
using NUnit.Framework;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ClosedXML.Tests.Excel.InsertData
{
    public class ObjectReaderTests
    {
        private static readonly TablesTests.TestObjectWithAttributes[] ObjectWithAttributes =
        {
            new()
            {
                Column1 = "Value 1",
                Column2 = "Value 2",
                UnOrderedColumn = 3,
                MyField = 4,
            },
            new()
            {
                Column1 = "Value 5",
                Column2 = "Value 6",
                UnOrderedColumn = 7,
                MyField = 8,
            }
        };

        private static readonly TablesTests.TestObjectWithoutAttributes[] ObjectWithoutAttributes =
        {
            new()
            {
                Column1 = "Value 9",
                Column2 = "Value 10"
            },
            new()
            {
                Column1 = "Value 11",
                Column2 = "Value 12"
            }
        };

        private static readonly TestPoint[] Structs =
        {
            new()
            {
                X = 1,
                Y = 2,
                Z = 3
            },
            new(),
        };

        private static readonly TestPoint?[] NullableStructs =
        {
            new TestPoint
            {
                X = 1,
                Y = 2,
                Z = 3
            },
            new TestPoint(),
            null
        };

        [TestCaseSource(nameof(ObjectSourceNames))]
        public string CanGetPropertyName<T>(IEnumerable<T> data, int propertyIndex)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return reader.GetPropertyName(propertyIndex);
        }

        private static IEnumerable<TestCaseData> ObjectSourceNames
        {
            get
            {
                IEnumerable data = ObjectWithoutAttributes;
                yield return new TestCaseData(data, 0).Returns("Column1");
                yield return new TestCaseData(data, 1).Returns("Column2");

                data = ObjectWithAttributes;
                yield return new TestCaseData(data, 0).Returns("FirstColumn");
                yield return new TestCaseData(data, 1).Returns("SecondColumn");
                yield return new TestCaseData(data, 2).Returns("SomeFieldNotProperty");
                yield return new TestCaseData(data, 3).Returns("UnOrderedColumn");

                data = Structs;
                yield return new TestCaseData(data, 0).Returns("X");
                yield return new TestCaseData(data, 1).Returns("Y");
                yield return new TestCaseData(data, 2).Returns("Z");

                data = NullableStructs;
                yield return new TestCaseData(data, 0).Returns("X");
                yield return new TestCaseData(data, 1).Returns("Y");
                yield return new TestCaseData(data, 2).Returns("Z");
            }
        }

        [TestCaseSource(nameof(PropertyCounts))]
        public int CanGetPropertiesCount(IEnumerable data)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return reader.GetPropertiesCount();
        }

        private static IEnumerable<TestCaseData> PropertyCounts
        {
            get
            {
                IEnumerable data = ObjectWithoutAttributes;
                yield return new TestCaseData(data).Returns(2);

                data = ObjectWithAttributes;
                yield return new TestCaseData(data).Returns(4);

                data = Structs;
                yield return new TestCaseData(data).Returns(3);

                data = NullableStructs;
                yield return new TestCaseData(data).Returns(3);
            }
        }

        [Test]
        public void CanGetRecordsCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(ObjectWithAttributes);
            Assert.That(reader.GetRecords().Count(), Is.EqualTo(2));
        }

        [Test]
        public void CanReadValues_FromObject()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(ObjectWithAttributes);
            var result = reader.GetRecords();

            var firstRecord = result.First().ToArray();
            var lastRecord = result.Last().ToArray();

            Assert.Multiple(() =>
            {
                Assert.That(firstRecord[0], Is.EqualTo("Value 2"));
                Assert.That(firstRecord[1], Is.EqualTo("Value 1"));
                Assert.That(firstRecord[2], Is.EqualTo(4));
                Assert.That(firstRecord[3], Is.EqualTo(3));

                Assert.That(lastRecord[0], Is.EqualTo("Value 6"));
                Assert.That(lastRecord[1], Is.EqualTo("Value 5"));
                Assert.That(lastRecord[2], Is.EqualTo(8));
                Assert.That(lastRecord[3], Is.EqualTo(7));
            });
        }

        [Test]
        public void CanReadValues_FromStruct()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(Structs);
            var result = reader.GetRecords();

            var firstRecord = result.First().ToArray();
            var lastRecord = result.Last().ToArray();

            Assert.Multiple(() =>
            {
                Assert.That(firstRecord[0], Is.EqualTo(1));
                Assert.That(firstRecord[1], Is.EqualTo(2));
                Assert.That(firstRecord[2], Is.EqualTo(3));

                Assert.That(lastRecord[0], Is.EqualTo(0));
                Assert.That(lastRecord[1], Is.EqualTo(0));
                Assert.That(lastRecord[2], Is.EqualTo(Blank.Value));
            });
        }

        [Test]
        public void CanReadValues_FromNullableStruct()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(NullableStructs);
            var result = reader.GetRecords();

            var firstRecord = result.First().ToArray();
            var lastRecord = result.Last().ToArray();

            Assert.Multiple(() =>
            {
                Assert.That(firstRecord[0], Is.EqualTo(1));
                Assert.That(firstRecord[1], Is.EqualTo(2));
                Assert.That(firstRecord[2], Is.EqualTo(3));

                Assert.That(lastRecord[0], Is.EqualTo(Blank.Value));
                Assert.That(lastRecord[1], Is.EqualTo(Blank.Value));
                Assert.That(lastRecord[2], Is.EqualTo(Blank.Value));
            });
        }

        [Test]
        public void IgnoresIndexers()
        {
            var data = new[] { new TestClassWithIndexer() };
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);

            Assert.Multiple(() =>
            {
                Assert.That(reader.GetPropertiesCount(), Is.EqualTo(1));
                Assert.That(reader.GetPropertyName(0), Is.EqualTo(nameof(TestClassWithIndexer.Value)));
            });
        }

        private record TestClassWithIndexer
        {
            public int Value => 0;
            public int this[int i] => 0;
        }

        private struct TestPoint
        {
            public double X { get; set; }
            public double Y { get; set; }
            public double? Z { get; set; }
        }
    }
}
