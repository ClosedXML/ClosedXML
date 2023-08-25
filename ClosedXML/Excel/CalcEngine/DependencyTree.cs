using System;
using System.Collections.Generic;

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
        private readonly Dictionary<XLCellFormula, CellFormulaDependencies> _dependencies = new();
        private readonly DependenciesVisitor _visitor;
        private readonly CalcEngine _engine;

        public DependencyTree(CalcEngine engine)
        {
            _visitor = new DependenciesVisitor();
            _engine = engine;
        }

        /// <summary>
        /// Add a formula to the dependency tree.
        /// </summary>
        /// <param name="sheet">Sheet where is the formula.</param>
        /// <param name="formula">The cell formula.</param>
        /// <returns>Added cell formula dependencies.</returns>
        /// <exception cref="ArgumentException">Formula already is in the tree.</exception>
        internal CellFormulaDependencies AddFormula(XLWorksheet sheet, XLCellFormula formula)
        {
            var ast = formula.GetAst(_engine);
            var context = new DependenciesContext(new XLSheetArea(sheet.Name, formula.Range));
            var rootReference = ast.AstRoot.Accept(context, _visitor);

            // If formula references are propagated to the root, make sure to add them.
            if (rootReference is not null)
                context.AddAreas(rootReference);

            _dependencies.Add(formula, context.Dependencies);
            return context.Dependencies;
        }
    }
}
