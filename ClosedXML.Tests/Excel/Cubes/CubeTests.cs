using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Cubes
{
    [TestFixture]
    public class CubeTests
    {
        [Test]
        public void CalLoadAndSaveCubeFromRange()
        {
            // Disable validation, because connection type for range is 102 and validator expects at most 8.
            TestHelper.LoadSaveAndCompare(@"Other\Cubes\CubeFromRange-Input.xlsx", @"Other\Cubes\CubeFromRange-Output.xlsx", validate: false);
        }
    }
}
