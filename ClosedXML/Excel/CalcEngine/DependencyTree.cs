﻿using System;
using System.Collections.Generic;
using System.Linq;
using RBush;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// <para>
    /// A dependency tree structure to hold all formulas of the workbook and reference
    /// objects they depend on. The key feature of dependency tree is to propagate
    /// dirty flag across formulas.
    /// </para>
    /// <para>
    /// When a data in a cell changes, all formulas that depend on it should be marked
    /// as dirty, but it is hard to find which cells are affected - that is what
    /// dependency tree does.
    /// </para>
    /// <para>
    /// Dependency tree must be updated, when structure of a workbook is updated:
    /// <list type="bullet">
    ///   <item>Sheet is added, renamed or deleted.</item>
    ///   <item>Name is added or deleted.</item>
    ///   <item>Table is resized, renamed, added or deleted.</item>
    /// </list>
    /// Any such action changes what cells formula depends on and
    /// the formula dependencies must be updated.
    /// </para>
    /// </summary>
    internal class DependencyTree
    {
        private readonly XLWorkbook _workbook;

        /// <summary>
        /// The source of the truth, a storage of formula dependencies. The dependency tree is
        /// constructed from this collection.
        /// </summary>
        private readonly Dictionary<XLCellFormula, FormulaDependencies> _dependencies = new();

        /// <summary>
        /// Visitor to extract precedents of formulas.
        /// </summary>
        private readonly DependenciesVisitor _visitor;

        /// <summary>
        /// Engine to get AST of not-yet parsed formulas.
        /// </summary>
        private readonly CalcEngine _engine;

        /// <summary>
        /// A dependency tree for each sheet (key is sheet name).
        /// </summary>
        private readonly Dictionary<string, SheetDependencyTree> _sheetTrees = new(XLHelper.SheetComparer);

        public DependencyTree(XLWorkbook workbook)
        {
            _workbook = workbook;
            _visitor = new DependenciesVisitor();
            _engine = workbook.CalcEngine;
        }

        internal bool IsEmpty => _sheetTrees.All(sheetTree => sheetTree.Value.IsEmpty) && _dependencies.Count == 0;

        /// <summary>
        /// Add a formula to the dependency tree.
        /// </summary>
        /// <param name="formulaArea">Area of a formula, for normal cells 1x1, for array can be larger.</param>
        /// <param name="formula">The cell formula.</param>
        /// <returns>Added cell formula dependencies.</returns>
        /// <exception cref="ArgumentException">Formula already is in the tree.</exception>
        internal FormulaDependencies AddFormula(XLSheetArea formulaArea, XLCellFormula formula)
        {
            var precedents = GetFormulaPrecedents(formulaArea, formula);

            _dependencies.Add(formula, precedents);

            foreach (var precedentArea in precedents.Areas)
            {
                // Add dependency to its sheet dependency tree
                if (!_sheetTrees.TryGetValue(precedentArea.Name, out var sheetTree))
                {
                    sheetTree = new SheetDependencyTree();
                    _sheetTrees.Add(precedentArea.Name, sheetTree);
                }

                var dependent = new Dependent(formulaArea, formula);
                sheetTree.AddDependent(precedentArea.Area, dependent);
            }

            return precedents;
        }

        /// <summary>
        /// Remove formula from the dependency tree.
        /// </summary>
        /// <param name="formula">Formula to remove.</param>
        internal void RemoveFormula(XLCellFormula formula)
        {
            if (!_dependencies.TryGetValue(formula, out var dependencies))
                return;

            _dependencies.Remove(formula);
            foreach (var precedentArea in dependencies.Areas)
            {
                if (!_sheetTrees.TryGetValue(precedentArea.Name, out var sheetTree))
                    throw new InvalidOperationException($"Dependency tree for sheet '{precedentArea.Name}' not found.");

                sheetTree.RemoveDependent(precedentArea.Area, formula);
            }
        }

        /// <summary>
        /// Mark all formulas that depend (directly or transitively) on the area as dirty.
        /// </summary>
        internal void MarkDirty(XLSheetArea dirtyArea)
        {
            // BFS vs DFS: Although the longest chain found in the wild is 1000
            // formulas long, attacker could supply malicious excel with recursion
            // leading to stack overflow => use queue even with extra allocation cost.
            var queue = new Queue<XLSheetArea>();
            queue.Enqueue(dirtyArea);
            while (queue.Count > 0)
            {
                var affectedArea = queue.Dequeue();
                var sheetTree = _sheetTrees[affectedArea.Name];
                foreach (var area in sheetTree.FindDependentsAreas(affectedArea.Area))
                {
                    foreach (var dependent in area.Dependents)
                    {
                        // Ensure we don't end up in an infinite cycle
                        if (dependent.IsDirty)
                            continue;

                        dependent.MarkDirty();
                        queue.Enqueue(dependent.FormulaArea);
                    }
                }
            }
        }

        private FormulaDependencies GetFormulaPrecedents(XLSheetArea formulaArea, XLCellFormula formula)
        {
            var ast = formula.GetAst(_engine);
            var context = new DependenciesContext(formulaArea, _workbook);
            var rootReference = ast.AstRoot.Accept(context, _visitor);

            // If formula references are propagated to the root, make sure to add them.
            if (rootReference is not null)
                context.AddAreas(rootReference);

            return context.Dependencies;
        }

        /// <summary>
        /// An area that is referred by formulas in different cells, i.e. it
        /// contains precedent cells for a formula. If anything in the area
        /// potentially changes, all dependents might also change.
        /// </summary>
        private class AreaDependents : ISpatialData
        {
            /// <summary>
            /// An area in a sheet that is used by formulas, converted to RBush envelope.
            /// All RBush <c>double</c> coordinates are whole numbers.
            /// </summary>
            private readonly Envelope _area;

            private readonly List<Dependent> _dependents;

            internal AreaDependents(in Envelope area, Dependent firstDependent)
            {
                _area = area;
                _dependents = new List<Dependent> { firstDependent };
            }

            /// <summary>
            /// The area in a sheet on which some formulas depend on.
            /// </summary>
            /// <example><c>SIN(A4)</c> depends on <c>A4:A4</c> area.</example>.
            public ref readonly Envelope Envelope => ref _area;

            /// <summary>
            /// List of formulas that depend on the range, always at least one.
            /// </summary>
            internal IReadOnlyList<Dependent> Dependents => _dependents;

            internal void AddDependent(Dependent dependent)
            {
                _dependents.Add(dependent);
            }

            internal void RemoveDependent(XLCellFormula formula)
            {
                for (var i = 0; i < _dependents.Count; ++i)
                {
                    var dependent = _dependents[i];

                    // several different formulas can depend on same area,
                    // remove only dependent of the formula.
                    if (dependent.Formula != formula)
                        continue;

                    // Remove from list by moving the last element to the removed
                    // element place and decrease capacity.
                    _dependents[i] = _dependents[_dependents.Count - 1];

                    // Remove last item, capacity is unchanged, only list size is updated.
                    _dependents.RemoveAt(_dependents.Count - 1);
                }
            }
        }

        /// <summary>
        /// A dependent on a precedent area. If the precedent area changes,
        /// the dependent might also now be invalid.
        /// </summary>
        private readonly struct Dependent
        {
            /// <summary>
            /// Area that is invalidated, when precedent area is marked as
            /// dirty. Generally, it is an area of formula (1x1 for normal
            /// formulas), larger for array formulas. Cell formula by itself
            /// doesn't contain it's address to make it easier add/delete
            /// rows/cols.
            /// </summary>
            internal readonly XLSheetArea FormulaArea;

            internal Dependent(XLSheetArea formulaArea, XLCellFormula formula)
            {
                FormulaArea = formulaArea;
                Formula = formula;
            }

            /// <summary>
            /// The formula that is affected by changes in precedent area.
            /// </summary>
            internal XLCellFormula Formula { get; }

            internal bool IsDirty => Formula.IsDirty;

            internal bool MarkDirty() => Formula.IsDirty = true;
        }

        /// <summary>
        /// A dependency tree for a single worksheet.
        /// </summary>
        private class SheetDependencyTree
        {
            /// <summary>
            /// The precedent areas are not duplicated, though two areas might overlap.
            /// </summary>
            private readonly RBush<AreaDependents> _tree;

            /// <summary>
            /// All precedent areas in the sheet for all formulas in the workbook.
            /// </summary>
            /// <remarks>
            /// Not sure extra memory (at least 32 bytes per formula) is worth less CPU: O(1) vs O(log N)....
            /// </remarks>
            private readonly Dictionary<XLSheetRange, AreaDependents> _precedentAreas;

            internal SheetDependencyTree()
            {
                _tree = new RBush<AreaDependents>();
                _precedentAreas = new Dictionary<XLSheetRange, AreaDependents>();
            }

            internal bool IsEmpty => _tree.Count == 0;

            internal void AddDependent(XLSheetRange precedentRange, Dependent dependent)
            {
                if (!_precedentAreas.TryGetValue(precedentRange, out var precedentArea))
                {
                    precedentArea = new AreaDependents(ToEnvelope(precedentRange), dependent);
                    _precedentAreas.Add(precedentRange, precedentArea);
                    _tree.Insert(precedentArea);
                }
                else
                {
                    precedentArea.AddDependent(dependent);
                }
            }

            internal IReadOnlyList<AreaDependents> FindDependentsAreas(XLSheetRange dirtyRange)
            {
                return _tree.Search(ToEnvelope(dirtyRange));
            }

            /// <summary>
            /// Remove a dependency of <paramref name="formula"/> on a
            /// <paramref name="precedentRange"/> from the sheet dependency tree.
            /// </summary>
            /// <param name="precedentRange">A precedent area in the sheet.</param>
            /// <param name="formula">Formula depending on the <paramref name="precedentRange"/>.</param>
            internal void RemoveDependent(XLSheetRange precedentRange, XLCellFormula formula)
            {
                if (!_precedentAreas.TryGetValue(precedentRange, out var precedentArea))
                    return;

                precedentArea.RemoveDependent(formula);
                if (precedentArea.Dependents.Count == 0)
                {
                    _tree.Delete(precedentArea);
                    _precedentAreas.Remove(precedentRange);
                }
            }

            private static Envelope ToEnvelope(XLSheetRange range)
            {
                return new Envelope(range.LeftColumn, range.TopRow, range.RightColumn, range.BottomRow);
            }
        }
    }
}
