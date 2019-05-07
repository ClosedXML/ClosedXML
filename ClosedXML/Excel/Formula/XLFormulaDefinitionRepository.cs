// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel.Caching;

namespace ClosedXML.Excel
{
    internal class XLFormulaDefinitionRepository : XLRepositoryBase<XLFormulaDefinitionKey, XLFormulaDefinition>
    {
        #region Public Constructors

        public XLFormulaDefinitionRepository(Func<XLFormulaDefinitionKey, XLFormulaDefinition> createNew)
            : base(createNew)
        {
        }

        public XLFormulaDefinitionRepository(Func<XLFormulaDefinitionKey, XLFormulaDefinition> createNew,
            IEqualityComparer<XLFormulaDefinitionKey> comparer)
            : base(createNew, comparer)
        {
        }

        #endregion Public Constructors

        public XLFormulaDefinition GetOrCreate(string formulaR1C1)
        {
            return GetOrCreate(new XLFormulaDefinitionKey(formulaR1C1));
        }

        public XLFormulaDefinition Store(XLFormulaDefinition formula)
        {
            if (formula == null) throw new ArgumentNullException(nameof(formula));

            return Store(new XLFormulaDefinitionKey(formula.GetFormulaR1C1()), formula);
        }
    }
}
