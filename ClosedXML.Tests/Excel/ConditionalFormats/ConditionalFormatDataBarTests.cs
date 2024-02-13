using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.ConditionalFormats
{
    [TestFixture]
    public class ConditionalFormatDataBarTests
    {
        [Test]
        public void ConditionalFormat_allows_databars_with_and_without_gradients()
        {
            TestHelper.CreateSaveLoadAssert(
                (wb, ws) =>
                {
                    ws.Cell("A1").Value = 0.25;
                    ws.Cell("A2").Value = 0.5;
                    ws.Cell("A3").Value = 0.75;

                    // Default (will use gradient)
                    ws.Cell("A1").AddConditionalFormat()
                      .AddDataBar(XLColor.Red)
                      .Minimum(XLCFContentType.Number, 0)
                      .Maximum(XLCFContentType.Number, 1);

                    // Solid fill
                    ws.Cell("A2").AddConditionalFormat()
                      .AddDataBar(XLColor.Red)
                      .Minimum(XLCFContentType.Number, 0)
                      .Maximum(XLCFContentType.Number, 1)
                      .SetGradient(false);

                    // Gradient
                    ws.Cell("A3").AddConditionalFormat()
                      .AddDataBar(XLColor.Red)
                      .Minimum(XLCFContentType.Number, 0)
                      .Maximum(XLCFContentType.Number, 1)
                      .SetGradient(true);
                },
                (_, ws) =>
                {
                    Assert.AreEqual(true, ws.ConditionalFormats.Single(cf => cf.Range == ws.Cell("A1").AsRange()).DataBar?.Gradient);
                    Assert.AreEqual(false, ws.ConditionalFormats.Single(cf => cf.Range == ws.Cell("A2").AsRange()).DataBar?.Gradient);
                    Assert.AreEqual(true, ws.ConditionalFormats.Single(cf => cf.Range == ws.Cell("A3").AsRange()).DataBar?.Gradient);
                },
                @"Examples\ConditionalFormatting\CFDataBarGradient.xlsx");
        }
    }
}
