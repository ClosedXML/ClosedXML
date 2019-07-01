// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A generic version of <see cref="XLContextualFormula"/> that may relate to XLCell, XLConditionalFormat, etc.
    /// </summary>
    internal abstract class XLContextualFormula<T> : XLContextualFormula where T : class
    {
        #region Public Properties

        public T Context { get; }

        #endregion Public Properties

        #region Public Constructors

        public XLContextualFormula(T context, XLFormulaDefinition formulaDefinition) : base(formulaDefinition)
        {
            Context = context ?? throw new ArgumentNullException(nameof(context));
        }

        #endregion Public Constructors
    }

    /// <summary>
    /// A base class for storing formula related to a certain context (XLCell, for example).
    /// The simpler name XLFormula has already been taken by another class publicly exposed,
    /// so don't get confused.
    /// </summary>
    internal abstract class XLContextualFormula
    {
        #region Public Properties

        public abstract XLAddress BaseAddress { get; }

        public string FormulaA1 => _formulaDefinition.GetFormulaA1(BaseAddress);

        public string FormulaR1C1 => _formulaDefinition.GetFormulaR1C1();

        #endregion Public Properties

        #region Protected Constructors

        protected XLContextualFormula(XLFormulaDefinition formulaDefinition)
        {
            _formulaDefinition = formulaDefinition ?? throw new ArgumentNullException(nameof(formulaDefinition));
        }

        #endregion Protected Constructors

        #region Private Fields

        private readonly XLFormulaDefinition _formulaDefinition;

        #endregion Private Fields
    }
}
