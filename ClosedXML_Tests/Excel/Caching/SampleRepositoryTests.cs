using System;
using ClosedXML.Excel;
using ClosedXML.Excel.Caching;
using NUnit.Framework;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ClosedXML_Tests.Excel.Caching
{
    [TestFixture]
    public class BaseRepositoryTests
    {
        [Test]
        public void DifferentEntitiesWithSameKeyStoredOnce()
        {
            // Arrange
            int key = 12345;
            var entity1 = new SampleEntity(key);
            var entity2 = new SampleEntity(key);
            var sampleRepository = this.CreateSampleRepository();

            // Act
            var storedEntity1 = sampleRepository.Store(key, entity1);
            var storedEntity2 = sampleRepository.Store(key, entity2);

            // Assert
            Assert.AreSame(entity1, storedEntity1);
            Assert.AreSame(entity1, storedEntity2);
            Assert.AreNotSame(entity2, storedEntity2);
        }


        [Test]
        public void NonUsedReferencesAreGCed()
        {
#if !DEBUG
            // Arrange
            int key = 12345;
            var sampleRepository = this.CreateSampleRepository();

            // Act
            var storedEntityRef1 = new WeakReference(sampleRepository.Store(key, new SampleEntity(key)));

            int count = 0;
            do
            {
                Thread.Sleep(50);
                GC.Collect();
                count++;
            } while (storedEntityRef1.IsAlive && count < 10);

            // Assert
            if (count == 10)
                Assert.Fail("storedEntityRef1 was not GCed");

            Assert.IsFalse(sampleRepository.Any());
#else
            Assert.Ignore("Can't run in DEBUG");
#endif
        }


        [Test]
        public void NonUsedReferencesAreGCed2()
        {
#if !DEBUG
            // Arrange
            int countUnique = 30;
            int repeatCount = 1000;
            SampleEntity[] entities = new SampleEntity[countUnique * repeatCount];
            for (int i = 0; i < countUnique; i++)
            {
                for (int j = 0; j < repeatCount; j++)
                {
                    entities[i * repeatCount + j] = new SampleEntity(i);
                }
            }

            var sampleRepository = this.CreateSampleRepository();

            // Act
            Parallel.ForEach(entities, new ParallelOptions { MaxDegreeOfParallelism = 8 },
                e => sampleRepository.Store(e.Key, e));

            Thread.Sleep(50);
            GC.Collect();
            var storedEntries = sampleRepository.ToList();

            // Assert
            Assert.AreEqual(0, storedEntries.Count);
#else
            Assert.Ignore("Can't run in DEBUG");
#endif
        }

        [Test]
        public void ConcurrentAddingCausesNoDuplication()
        {
            // Arrange
            int countUnique = 30;
            int repeatCount = 1000;
            SampleEntity[] entities = new SampleEntity[countUnique * repeatCount];
            for (int i = 0; i < countUnique; i++)
            {
                for (int j = 0; j < repeatCount; j++)
                {
                    entities[i * repeatCount + j] = new SampleEntity(i);
                }
            }

            var sampleRepository = this.CreateSampleRepository();

            // Act
            Parallel.ForEach(entities, new ParallelOptions { MaxDegreeOfParallelism = 8 },
                e => sampleRepository.Store(e.Key, e));
            var storedEntries = sampleRepository.ToList();

            // Assert
            Assert.AreEqual(countUnique, storedEntries.Count);
            Assert.NotNull(entities); // To protect them from GC
        }

        [Test]
        public void ReplaceKeyInRepository()
        {
            // Arrange
            int key1 = 12345;
            int key2 = 54321;
            var entity = new SampleEntity(key1);
            var sampleRepository = this.CreateSampleRepository();
            var storedEntity1 = sampleRepository.Store(key1, entity);

            // Act
            sampleRepository.Replace(key1, key2);
            bool containsOld = sampleRepository.ContainsKey(key1, out var _);
            bool containsNew = sampleRepository.ContainsKey(key2, out var _);
            var storedEntity2 = sampleRepository.GetOrCreate(key2);

            // Assert
            Assert.IsFalse(containsOld);
            Assert.IsTrue(containsNew);
            Assert.AreSame(entity, storedEntity1);
            Assert.AreSame(entity, storedEntity2);
        }

        [Test]
        public void ConcurrentReplaceKeyInRepository()
        {
            var sampleRepository = new EditableRepository();
            int[] keys = Enumerable.Range(0, 1000).ToArray();
            keys.ForEach(key => sampleRepository.GetOrCreate(key));

            Parallel.ForEach(keys, key =>
            {
                var val1 = sampleRepository.Replace(key, key + 2000);
                val1.Key = key + 2000;
                var val2 = sampleRepository.GetOrCreate(key + 2000);
                Assert.AreSame(val1, val2);
            });
        }

        [Test]
        public void ReplaceNonExistingKeyInRepository()
        {
            int key1 = 100;
            int key2 = 200;
            int key3 = 300;
            var entity = new SampleEntity(key1);
            var sampleRepository = this.CreateSampleRepository();
            sampleRepository.Store(key1, entity);

            sampleRepository.Replace(key2, key3);
            var all = sampleRepository.ToList();

            Assert.AreEqual(1, all.Count);
            Assert.AreSame(entity, all.First());
        }

        private SampleRepository CreateSampleRepository()
        {
            return new SampleRepository();
        }

        /// <summary>
        /// Class under testing
        /// </summary>
        internal class SampleRepository : XLRepositoryBase<int, SampleEntity>
        {
            public SampleRepository() : base(key => new SampleEntity(key))
            {
            }
        }

        public class SampleEntity
        {
            public int Key { get; private set; }

            public SampleEntity(int key)
            {
                Key = key;
            }
        }


        /// <summary>
        /// Class under testing
        /// </summary>
        internal class EditableRepository : XLRepositoryBase<int, EditableEntity>
        {
            public EditableRepository() : base(key => new EditableEntity(key))
            {
            }
        }

        public class EditableEntity
        {
            public int Key { get; set; }

            public EditableEntity(int key)
            {
                Key = key;
            }
        }
    }


}
