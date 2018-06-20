using ClosedXML_Examples.Sparklines;
using NUnit.Framework;

namespace ClosedXML_Tests.Examples
{
    [TestFixture]
    public class SparklineTests
    {
        [Test]
        public void AddingSparklines()
        {
            TestHelper.RunTestExample<AddingSparklines>(@"Sparklines\AddingSparklines.xlsx");
        }

        [Test]
        public void DeletingSparklines()
        {
            TestHelper.RunTestExample<DeletingSparklines>(@"Sparklines\DeletingSparklines.xlsx");
        }

        [Test]
        public void CopyingSparklines()
        {
            TestHelper.RunTestExample<CopyingSparklines>(@"Sparklines\CopyingSparklines.xlsx");
        }

        [Test]
        public void DeletingSparklinesHelp()
        {
            TestHelper.RunTestExample<DeletingSparklinesHelp>(@"Sparklines\DeletingSparklinesHelp.xlsx");
        }

    }
}
