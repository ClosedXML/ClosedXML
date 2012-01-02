using ClosedXML_Examples;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Examples
{
    [TestClass]
    public class CommentsTests
    {
    
    [TestMethod]
    public void AddingComments()
    {
        TestHelper.RunTestExample<AddingComments>(@"Comments\AddingComments.xlsx");
    }
    
    }
}