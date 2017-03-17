using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML_Tests
{
    [TestFixture]

    public class RemoveDeprecatedCodeTests
    {
        [Test]
        public void MakeSureObsoleteAttributeGetsRemovedBeforeV1_0()
        {

            ///////////////////////////////////////////////////
            // This test is just a reminder to remove the    //
            // obsolete ColumnOrderAttribute after a         //
            // reasonable grace period.                      //
            // Remove the ColumnOrderAttribute and this test //
            //                                               //
            // Sorry dear future developer if your were      //
            // scared of confused                            //
            ///////////////////////////////////////////////////


#pragma warning disable 612, 618 // Make the test compile without warning even though the attribute is obsolete
            var type = typeof(ClosedXML.Attributes.ColumnOrderAttribute);
#pragma warning restore 612, 618
            var assembly = type.Assembly;
            var version = assembly.GetCustomAttributes(typeof(System.Reflection.AssemblyFileVersionAttribute), false).OfType<System.Reflection.AssemblyFileVersionAttribute>().Single();
            Assert.AreEqual('0', version.Version[0], "Once ClosedXML reaches v1.0 the ColumnOrderAttribute should be removed"); // Break tests when the first version digit is 1 or higher
        }
    }
}
