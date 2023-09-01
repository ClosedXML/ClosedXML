using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    internal class DependencyTreeTests
    {
        [Test]
        [TestCaseSource(nameof(AreaDependenciesTestCases))]
        public void Area_dependencies_are_extracted_from_formula(string formula, IReadOnlyList<XLSheetArea> expectedAreas)
        {
            var dependencies = GetDependencies(formula);
            CollectionAssert.AreEquivalent(expectedAreas, dependencies.Areas);
        }

        [Test]
        [TestCaseSource(nameof(NameDependenciesTestCases))]
        public void Name_dependencies_are_kept_for_dependencies_update(string formula, IReadOnlyList<XLName> expectedNames)
        {
            var dependencies = GetDependencies(formula);
            CollectionAssert.AreEquivalent(expectedNames, dependencies.Names);
        }

        [Test]
        public void Name_range_is_added_to_dependencies_of_formula()
        {
            var dependencies = GetDependencies("name + D2", init: wb =>
            {
                wb.NamedRanges.Add("name", "Sheet!$B$4+Sheet!$C$6");
            });
            CollectionAssert.AreEquivalent(new XLSheetArea[]
            {
                new("Sheet", XLSheetRange.Parse("D2")),
                new("Sheet", XLSheetRange.Parse("B4")),
                new("Sheet", XLSheetRange.Parse("C6"))
            }, dependencies.Areas);
            CollectionAssert.AreEquivalent(new[] { new XLName("name") }, dependencies.Names);
        }

        [Test]
        public void Name_range_that_is_reference_is_propagated_to_formula()
        {
            var dependencies = GetDependencies("B3:name", init: wb =>
            {
                wb.NamedRanges.Add("name", "Sheet!$D$7");
            });
            CollectionAssert.AreEquivalent(new XLSheetArea[]
            {
                new("Sheet", XLSheetRange.Parse("B3:D7")),
            }, dependencies.Areas);
            CollectionAssert.AreEquivalent(new[] { new XLName("name") }, dependencies.Names);
        }

        [Test]
        public void Name_range_can_used_another_name_range()
        {
            var dependencies = GetDependencies("outer", init: wb =>
            {
                wb.NamedRanges.Add("outer", "Sheet!$D$7 + inner");
                wb.NamedRanges.Add("inner", "Sheet!$B$1");
            });
            CollectionAssert.AreEquivalent(new XLSheetArea[]
            {
                new("Sheet", XLSheetRange.Parse("D7")),
                new("Sheet", XLSheetRange.Parse("B1")),
            }, dependencies.Areas);
            CollectionAssert.AreEquivalent(new[] { new XLName("outer"), new XLName("inner") }, dependencies.Names);
        }

        [Test]
        public void Name_range_that_is_not_a_reference_can_be_added_to_dependency_tree_without_exception()
        {
            var dependencies = GetDependencies("name", init: wb =>
            {
                wb.NamedRanges.Add("name", "1+3");
            });
            CollectionAssert.IsEmpty(dependencies.Areas);
            CollectionAssert.AreEquivalent(new[] { new XLName("name") }, dependencies.Names);
        }

        [Test]
        public void Name_range_can_be_sheet_scoped_even_without_specified_sheet()
        {
            // Formula that references a name that is ambiguous between workbook and worksheet scoped one.
            const string formula = "name";
            var dependencies = GetDependencies(formula, init: wb =>
            {
                // Define two names, the local one should be selected
                wb.Worksheet("Sheet").NamedRanges.Add("name", "Sheet!$A$1");
                wb.NamedRanges.Add("name", "Sheet!$B$10");
            });
            CollectionAssert.AreEquivalent(new XLSheetArea[]
            {
                new("Sheet", XLSheetRange.Parse("A1"))
            }, dependencies.Areas);
            CollectionAssert.AreEquivalent(new[] { new XLName("name") }, dependencies.Names);
        }

        [Test]
        [Ignore("A1 to R1C1 conversion not yet implemented and the name formula must be parsed")]
        public void Name_range_that_uses_relative_reference_determines_actual_precedent_areas_through_cell_location()
        {
            var dependencies = GetDependencies("name", "D8", init: wb =>
            {
                wb.NamedRanges.Add("name", "Sheet!B4"); // equivalent of R[3]C[2]
            });
            CollectionAssert.AreEquivalent(new XLSheetArea[]
            {
                new("Sheet", XLSheetRange.Parse("F7")), // D4 (formula cell) + R[3]C[2] (name relative reference) = F7
            }, dependencies.Areas);
            CollectionAssert.AreEquivalent(new[] { new XLName("name") }, dependencies.Names);
        }

        #region Mark dirty

        [Test]
        public void Mark_dirty_single_chain_is_fully_marked()
        {
            using var wb = new XLWorkbook();
            var tree = new DependencyTree(wb);
            var ws = wb.AddWorksheet();
            AddFormula(tree, ws, "A2", "=A1");
            AddFormula(tree, ws, "A3", "=A2");
            AddFormula(tree, ws, "A4", "=A3");

            MarkDirty(tree, ws, "A1");
            AssertDirty(ws, "A2", "A3", "A4");
        }

        [Test]
        public void Mark_dirty_split_and_join_is_fully_marked()
        {
            using var wb = new XLWorkbook();
            var tree = new DependencyTree(wb);
            var ws1 = wb.AddWorksheet();
            AddFormula(tree, ws1, "B2", "=B1");
            AddFormula(tree, ws1, "C1", "=B2");
            AddFormula(tree, ws1, "C3", "=B2");
            AddFormula(tree, ws1, "D2", "=C1 + C3");

            MarkDirty(tree, ws1, "B1");
            AssertDirty(ws1, "B2", "C1", "C3", "D2");
        }

        [Test]
        public void Mark_dirty_uses_correct_sheet()
        {
            using var wb = new XLWorkbook();
            var tree = new DependencyTree(wb);
            var ws1 = wb.AddWorksheet("Sheet1");
            var ws2 = wb.AddWorksheet("Sheet2");

            // Make a chain, where each cell is on an opposite sheet
            AddFormula(tree, ws1, "B1", "=Sheet2!A1");
            AddFormula(tree, ws2, "C1", "=Sheet1!B1");
            AddFormula(tree, ws1, "D1", "=Sheet2!C1");
            AddFormula(tree, ws2, "E1", "=Sheet1!D1");

            // Formulas on opposite sheet
            AddFormula(tree, ws2, "B1", "=Sheet1!A1");
            AddFormula(tree, ws1, "C1", "=Sheet2!B1");
            AddFormula(tree, ws2, "D1", "=Sheet1!C1");
            AddFormula(tree, ws1, "E1", "=Sheet2!D1");

            MarkDirty(tree, ws2, "A1");
            AssertDirty(ws1, "B1", "D1");
            AssertDirty(ws2, "C1", "E1");

            AssertNotDirty(ws1, "C1", "E1");
            AssertNotDirty(ws2, "B1", "D1");
        }

        [Test]
        public void Mark_dirty_stops_at_dirty_cell()
        {
            using var wb = new XLWorkbook();
            var tree = new DependencyTree(wb);
            var ws = wb.AddWorksheet();
            AddFormula(tree, ws, "A2", "=A1");
            AddFormula(tree, ws, "A3", "=A2");
            AddFormula(tree, ws, "A4", "=A3");

            // Mark the middle one dirty, but A4 is still clear
            ((XLCell)ws.Cell("A3")).Formula.IsDirty = true;

            MarkDirty(tree, ws, "A1");
            AssertDirty(ws, "A2", "A3");
            AssertNotDirty(ws, "A4"); // Propagation stopped at the dirty cell A3.
        }

        [Test]
        public void Mark_dirty_wont_crash_on_cycle()
        {
            using var wb = new XLWorkbook();
            var tree = new DependencyTree(wb);
            var ws = wb.AddWorksheet();
            AddFormula(tree, ws, "B1", "=D1 + A1");
            AddFormula(tree, ws, "C1", "=B1");
            AddFormula(tree, ws, "D1", "=C1");

            // Tail depending on the cycle
            AddFormula(tree, ws, "E1", "=D1");

            MarkDirty(tree, ws, "A1");
            AssertDirty(ws, "B1", "C1", "D1", "E1");
        }

        [Test]
        public void Mark_dirty_affects_precedents_with_partial_overlap()
        {
            using var wb = new XLWorkbook();
            var tree = new DependencyTree(wb);
            var ws = wb.AddWorksheet();
            AddFormula(tree, ws, "D1", "=A1:B3");

            // B3:D4 overlaps with A1:B3 in B3
            MarkDirty(tree, ws, "B3:D4");
            AssertDirty(ws, "D1");
        }

        [Test]
        public void Mark_dirty_can_affect_multiple_chains_at_once()
        {
            using var wb = new XLWorkbook();
            var tree = new DependencyTree(wb);
            var ws = wb.AddWorksheet();
            AddFormula(tree, ws, "B1", "=A1");
            AddFormula(tree, ws, "B2", "=A2");
            AddFormula(tree, ws, "B3", "=A3");

            MarkDirty(tree, ws, "A2:A3");
            AssertDirty(ws, "B2", "B3");
            AssertNotDirty(ws, "B1");
        }

        #endregion

        private static void AddFormula(DependencyTree tree, IXLWorksheet sheet, string address, string formula)
        {
            var cell = (XLCell)sheet.Cell(address);
            cell.FormulaA1 = formula;
            var cellArea = new XLSheetArea(sheet.Name, new XLSheetRange(cell.SheetPoint, cell.SheetPoint));
            tree.AddFormula(cellArea, cell.Formula);
        }

        private static void MarkDirty(DependencyTree tree, IXLWorksheet sheet, string range)
        {
            var area = new XLSheetArea(sheet.Name, XLSheetRange.Parse(range));
            tree.MarkDirty(area);
        }

        private static void AssertDirty(IXLWorksheet sheet, params string[] dirtyRanges)
        {
            AssertDirtyFlag(true, sheet, dirtyRanges);
        }
        private static void AssertNotDirty(IXLWorksheet sheet, params string[] dirtyRanges)
        {
            AssertDirtyFlag(false, sheet, dirtyRanges);
        }

        private static void AssertDirtyFlag(bool expectedDirtyFlag, IXLWorksheet sheet, params string[] dirtyRanges)
        {
            var ws = (XLWorksheet)sheet;
            foreach (var dirtyRange in dirtyRanges)
            {
                foreach (var dirtyCell in ws.Cells(dirtyRange))
                {
                    Assert.AreEqual(expectedDirtyFlag, dirtyCell.Formula?.IsDirty);
                }
            }
        }

        private static FormulaDependencies GetDependencies(string formula, string formulaAddress = "A1", Action<XLWorkbook> init = null)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            init?.Invoke(wb);
            var tree = new DependencyTree(wb);
            var cell = ws.Cell(formulaAddress);
            cell.SetFormulaA1(formula);

            var cellFormula = ((XLCell)cell).Formula;
            var dependencies = tree.AddFormula(new XLSheetArea(ws.Name, cellFormula.Range), cellFormula);
            return dependencies;
        }

        public static IEnumerable<object[]> AreaDependenciesTestCases
        {
            get
            {
                // When a visitor visits a node, there are two choices for found references:
                // * propagate the reference to parent node (in most cases checked by range operator)
                // * add the reference directly to the dependencies

                // A formula that is a simple reference is propagated to the root
                yield return new object[]
                {
                    "A1",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1"))
                    }
                };

                // References are in a multiple levels of an expression without ref expression or
                // a function are added
                yield return new object[]
                {
                    "7+A1/(B1+C1)",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1")),
                        new XLSheetArea("Sheet", XLSheetRange.Parse("B1")),
                        new XLSheetArea("Sheet", XLSheetRange.Parse("C1"))
                    }
                };

                // Unary implicit intersection is propagated
                yield return new object[]
                {
                    // Due to issue ClosedParser#1, implicit intersection is not a part
                    // of ref_expression and I can't use `D3:@A1:C2` as a test case
                    "@A1:A4",
                    new[]
                    {
                        // Implicit intersection 
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1:A4")),
                    }
                };

                // Unary spill operator propagates a reference
                yield return new object[]
                {
                    "F2#:A7",
                    new[]
                    {
                        // This is not correct, but until spill operator works,
                        // but for now it provides best approximate for now.
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A2:F7")),
                    }
                };

                // Unary value operators (in this case percent) applied on reference adds it
                yield return new object[]
                {
                    "4+A4%",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A4")),
                    }
                };

                // Union operation propagates references
                yield return new object[]
                {
                    "(A1:B2,C1:D2):E3",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1:E3"))
                    }
                };

                // Range operation propagates
                yield return new object[]
                {
                    // Due to greedy nature, the A1:C4 is the first reference and D2 is the second
                    "A1:C4:D2",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1:D4")),
                    }
                };

                // Range operation with multiple operands
                yield return new object[]
                {
                    "A1:C4:IF(E10, D2, A10)",
                    new[]
                    {
                        // E10 is a value argument, thus isn't propagated, only added
                        new XLSheetArea("Sheet", XLSheetRange.Parse("E10")),
                        // Areas from same sheet are unified into a single larger area
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1:D10"))
                    }
                };

                // Range operator with multiple combinations
                yield return new object[]
                {
                    "IF(G4,Sheet!A1,Other!A2):IF(H3,Other!C4,C5)",
                    new[]
                    {
                        // G4 and H3 are not propagated to range operation, only added
                        new XLSheetArea("Sheet", XLSheetRange.Parse("G4")),
                        new XLSheetArea("Sheet", XLSheetRange.Parse("H3")),

                        // Largest possible area in each sheet, based on references in the sheet
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1:C5")),
                        new XLSheetArea("Other", XLSheetRange.Parse("A2:C4"))
                    }
                };

                // Range operation when an argument isn't a reference doesn't
                // create a range from both, adds
                yield return new object[]
                {
                    "INDEX({1},1,1):D2",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("D2")),
                    }
                };

                // Intersection - special case of one area against another area
                yield return new object[]
                {
                    "A1:C3 B2:D2",
                    new[]
                    {
                        // In this special case, intersection is evaluated
                        new XLSheetArea("Sheet", XLSheetRange.Parse("B2:C2")),
                    }
                };

                // Intersection - multi area operands. Due to complexity, keep
                // original ranges as dependencies.
                yield return new object[]
                {
                    "A1:E10 IF(TRUE,A1:C3,B2:D2)",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1:C3")),
                        new XLSheetArea("Sheet", XLSheetRange.Parse("B2:D2")),
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1:E10")),
                    }
                };

                // Value binary operation on references adds the references
                yield return new object[]
                {
                    "A1:B2 + A1:C4",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1:B2")),
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1:C4")),
                    }
                };

                // IF function - value is added and true/false values are propagated
                yield return new object[]
                {
                    "IF(A1,B1,C1):D2",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1")),
                        new XLSheetArea("Sheet", XLSheetRange.Parse("B1:D2")),
                    }
                };

                // IF function, but only false argument is reference
                yield return new object[]
                {
                    "IF(A1,5,B1):D2",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1")),
                        new XLSheetArea("Sheet", XLSheetRange.Parse("B1:D2")),
                    }
                };

                // IF function, but only true argument is reference and is propagated
                yield return new object[]
                {
                    "IF(A1,B1):D2",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1")),
                        new XLSheetArea("Sheet", XLSheetRange.Parse("B1:D2")),
                    }
                };

                // INDEX function propagates whole range of first argument
                yield return new object[]
                {
                    "INDEX(A1:C4,2,5):D2",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1:D4")),
                    }
                };

                // CHOOSE function adds first argument and propagates remaining arguments
                yield return new object[]
                {
                    "CHOOSE(A1,B1,5,C1):D2",
                    new[]
                    {
                        new XLSheetArea("Sheet", XLSheetRange.Parse("A1")),
                        new XLSheetArea("Sheet", XLSheetRange.Parse("B1:D2")),
                    }
                };

                // Non-ref functions add arguments
                yield return new object[]
                {
                    "POWER(SomeSheet!C4,Other!B1)",
                    new[]
                    {
                        new XLSheetArea("SomeSheet", XLSheetRange.Parse("C4")),
                        new XLSheetArea("Other", XLSheetRange.Parse("B1")),
                    }
                };
            }
        }

        public static IEnumerable<object[]> NameDependenciesTestCases
        {
            get
            {
                yield return new object[]
                {
                    "WorkbookName  + 5",
                    new[] { new XLName("WorkbookName") }
                };

                yield return new object[]
                {
                    "Sheet!Name",
                    new[] { new XLName("Sheet", "Name") }
                };
            }
        }
    }
}
