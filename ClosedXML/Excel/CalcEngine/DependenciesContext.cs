using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Context for <see cref="DependenciesVisitor"/>, it is used
    /// to collect all objects a formula depends on during calculation.
    /// </summary>
    internal class DependenciesContext
    {
        internal DependenciesContext(XLBookArea formulaArea, XLWorkbook workbook)
        {
            FormulaArea = formulaArea;
            Workbook = workbook;
        }

        /// <summary>
        /// An area of a formula, in most cases just one cell, for array formulas area of cells.
        /// </summary>
        internal XLBookArea FormulaArea { get; }

        internal XLWorkbook Workbook { get; }

        /// <summary>
        /// The result. Visitor adds all areas/names formula depends on to this.
        /// </summary>
        internal FormulaDependencies Dependencies { get; } = new();

        /// <summary>
        /// Add areas to a list of areas the formula depends on. Disregards duplicate entries.
        /// </summary>
        internal void AddAreas(List<XLBookArea> sheetAreas) => Dependencies.AddAreas(sheetAreas);

        /// <summary>
        /// Add name to a list of names the formula depends on. Disregards duplicate entries.
        /// </summary>
        internal void AddName(XLName name) => Dependencies.AddName(name);
    }
}
