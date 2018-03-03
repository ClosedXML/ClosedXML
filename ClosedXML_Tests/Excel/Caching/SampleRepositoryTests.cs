using ClosedXML.Excel.Caching;
using NUnit.Framework;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;

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
    }


}
