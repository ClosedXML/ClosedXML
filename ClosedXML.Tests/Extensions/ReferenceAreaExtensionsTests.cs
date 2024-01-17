using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Extensions;
using ClosedXML.Parser;
using NUnit.Framework;
using static ClosedXML.Parser.ReferenceAxisType;
using static ClosedXML.Parser.ReferenceStyle;

namespace ClosedXML.Tests.Extensions
{
    [TestFixture]
    internal class ReferenceAreaExtensionsTests
    {
        [Test]
        [TestCaseSource(nameof(A1TestCases))]
        public void ToSheetPoint_converts_a1_reference_to_sheet_range(ReferenceArea tokenArea, XLSheetRange expectedRange)
        {
            Assert.AreEqual(expectedRange, tokenArea.ToSheetRange(default));
        }

        [Test]
        [TestCaseSource(nameof(R1C1TestCases))]
        public void ToSheetPoint_converts_r1c1_reference_to_sheet_range(XLSheetPoint anchor, ReferenceArea tokenArea, XLSheetRange expectedRange)
        {
            Assert.AreEqual(expectedRange, tokenArea.ToSheetRange(anchor));
        }

        public static IEnumerable<object[]> A1TestCases()
        {
            // C5
            yield return new object[]
            {
                new ReferenceArea(Relative, 5, Relative, 3, A1),
                new XLSheetRange(5, 3, 5, 3)
            };

            // C5:E14
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 5, Relative, 3, A1), new RowCol(Relative, 14, Relative, 5, A1)),
                new XLSheetRange(5, 3, 14, 5)
            };

            // $B3:E$10
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 3, Absolute, 2, A1), new RowCol(Absolute, 10, Relative, 5, A1)),
                new XLSheetRange(3, 2, 10, 5)
            };

            // $B$3:$E$10
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Absolute, 3, Absolute, 2, A1), new RowCol(Absolute, 10, Absolute, 5, A1)),
                new XLSheetRange(3, 2, 10, 5)
            };

            // B10:E3 points are not in left top corner and bottom right corner
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 10, Relative, 2, A1), new RowCol(Absolute, 3, Absolute, 5, A1)),
                new XLSheetRange(3, 2, 10, 5)
            };

            // C:E
            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Relative, 3, A1), new RowCol(None, 0, Relative, 5, A1)),
                new XLSheetRange(XLHelper.MinRowNumber, 3, XLHelper.MaxRowNumber, 5)
            };

            // E:C
            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Relative, 5, A1), new RowCol(None, 0, Relative, 3, A1)),
                new XLSheetRange(XLHelper.MinRowNumber, 3, XLHelper.MaxRowNumber, 5)
            };

            // 14:30
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 14, None, 0, A1), new RowCol(Relative, 30, None, 0, A1)),
                new XLSheetRange(14, XLHelper.MinColumnNumber, 30, XLHelper.MaxColumnNumber)
            };

            // 30:14
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 30, None, 0, A1), new RowCol(Relative, 14, None, 0, A1)),
                new XLSheetRange(14, XLHelper.MinColumnNumber, 30, XLHelper.MaxColumnNumber)
            };
        }

        public static IEnumerable<object[]> R1C1TestCases()
        {
            // R2C4
            yield return new object[]
            {
                new XLSheetPoint(1, 1),
                new ReferenceArea(Absolute, 2, Absolute, 4, R1C1),
                new XLSheetRange(2, 4, 2, 4)
            };

            // R[2]C[4]
            yield return new object[]
            {
                new XLSheetPoint(3, 2), // R3C2
                new ReferenceArea(Relative, 2, Relative, 4, R1C1), // R[2]C[4]
                new XLSheetRange(5, 6, 5, 6)
            };

            // R[0]C[0] is the identical address
            yield return new object[]
            {
                new XLSheetPoint(3, 2), // R3C2
                new ReferenceArea(Relative, 0, Relative, 0, R1C1), // R[0]C[0]
                new XLSheetRange(3, 2, 3, 2)
            };

            // No looping: Maximum allowed value for relative column is `XLHelper.MaxColumnNumber-1`.
            yield return new object[]
            {
                new XLSheetPoint(1, 1), // R1C1
                new ReferenceArea(Relative, 0, Relative, 16383, R1C1), // R[0]C[16383]
                new XLSheetRange(1, XLHelper.MaxColumnNumber, 1, XLHelper.MaxColumnNumber)
            };

            // No looping: Minimum allowed value for relative column is `-XLHelper.MaxColumnNumber+1`.
            yield return new object[]
            {
                new XLSheetPoint(1, XLHelper.MaxColumnNumber), // R1C16384
                new ReferenceArea(Relative, 0, Relative, -16383, R1C1), // R[0]C[-16383]
                new XLSheetRange(1, 1, 1, 1) // R1C1
            };

            // Looping: when relative column adjusted to anchor is above the max column, it loops back
            yield return new object[]
            {
                new XLSheetPoint(1, 16380), // R1C16380
                new ReferenceArea(Relative, 0, Relative, 16380, R1C1), // R[0]C[16380]
                new XLSheetRange(1, 16376, 1, 16376) // RC16376
            };

            // Looping: when relative column adjusted to anchor is below the column 1, it loops back
            yield return new object[]
            {
                new XLSheetPoint(1, 10), // R1C10
                new ReferenceArea(Relative, 0, Relative, -16370, R1C1), // R[0]C[16370]
                new XLSheetRange(1, 24, 1, 24) // R1C24
            };

            // Looping: when relative row adjusted to anchor is above the max row, it loops back
            yield return new object[]
            {
                new XLSheetPoint(15, 1), // R15C1
                new ReferenceArea(Relative, 1048570, Relative, 0, R1C1), // R[1048570]C[0]
                new XLSheetRange(9, 1, 9, 1) // R9C1
            };

            // Looping: when relative row adjusted to anchor is below the row 1, it loops back
            yield return new object[]
            {
                new XLSheetPoint(1048570, 1), // R1048570C1
                new ReferenceArea(Relative, -1048573, Relative, 0, R1C1), // R[-1048573]C[0]
                new XLSheetRange(1048573, 1, 1048573, 1) // R1048573C1
            };

            // Area absolute
            yield return new object[]
            {
                new XLSheetPoint(754, 5742),
                new ReferenceArea(new RowCol(Absolute, 3, Absolute, 2, R1C1), new RowCol(Absolute, 7, Absolute, 4, R1C1)),
                XLSheetRange.Parse("B3:D7")
            };

            // Area relative
            yield return new object[]
            {
                new XLSheetPoint(3, 6),
                new ReferenceArea(new RowCol(Relative, 4, Relative, -1, R1C1), new RowCol(Relative, 6, Relative, 3, R1C1)), // R[4]C[-1]:R[6]C[3]
                new XLSheetRange(7, 5, 9, 9)
            };

            // Are with corners not in top left and right bottom
            yield return new object[]
            {
                new XLSheetPoint(3, 6),
                new ReferenceArea(new RowCol(Relative, 6, Relative, -1, R1C1), new RowCol(Relative, 4, Relative, 3, R1C1)), // R[6]C[-1]:R[4]C[3]
                new XLSheetRange(7, 5, 9, 9)
            };
        }
    }
}
