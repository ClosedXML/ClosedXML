// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A formula related to a particular cell.
    /// </summary>
    internal class XLCellFormula : XLContextualFormula<XLCell>
    {
        #region Public Constructors

        public XLCellFormula(XLCell cell, XLFormulaDefinition formulaDefinition)
            : base(cell, formulaDefinition)
        {
        }

        #endregion Public Constructors

        #region Public Methods

        public static XLCellFormula FromFormulaA1(XLCell cell, string formulaA1)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));

            var repository = cell.Worksheet?.Workbook?.FormulaDefinitionRepository;
            if (repository == null)
                throw new InvalidOperationException("Could not get FormulaDefinitionRepository");

            var formulaDefinition = new XLFormulaDefinition(formulaA1, cell.Address);
            formulaDefinition = repository.Store(formulaDefinition);

            return new XLCellFormula(cell, formulaDefinition);
        }

        public static XLCellFormula FromFormulaR1C1(XLCell cell, string formulaR1C1)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));

            var repository = cell.Worksheet?.Workbook?.FormulaDefinitionRepository;
            if (repository == null)
                throw new InvalidOperationException("Could not get FormulaDefinitionRepository");

            var formulaDefinition = new XLFormulaDefinition(formulaR1C1);
            formulaDefinition = repository.Store(formulaDefinition);

            return new XLCellFormula(cell, formulaDefinition);
        }

        #endregion Public Methods

        #region Overrides of XLContextualFormula

        public override XLAddress BaseAddress => Context.Address;

        #endregion Overrides of XLContextualFormula
    }
}
