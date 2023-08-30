using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A list of objects a cell formula depends on. If one of them changes,
    /// the formula value might no longer be accurate and needs to be recalculated.
    /// </summary>
    internal class FormulaDependencies
    {
        private readonly HashSet<XLSheetArea> _areas = new();
        private readonly HashSet<XLName> _names = new();

        /// <summary>
        /// List of areas the formula depends on. It is likely a superset of accurate
        /// result for unusual formulas, but if a value in an areas changes, the dependent
        /// formula should be marked as dirty.
        /// </summary>
        public IReadOnlyCollection<XLSheetArea> Areas => _areas;

        /// <summary>
        /// A collection of names in the formula. If a name changes (added, deleted),
        /// the formula dependencies should be refreshed, because new name might refer to
        /// different references (e.g. a name previously referred to <c>A5</c> and is redefined
        /// to <c>B7</c> or just value <c>7</c> =&gt; formula no longer depends on <c>A5</c>).
        /// </summary>
        public IReadOnlyCollection<XLName> Names => _names;
        
        internal void AddAreas(List<XLSheetArea> sheetAreas)
        {
            _areas.UnionWith(sheetAreas);
        }

        internal void AddName(XLName name)
        {
            _names.Add(name);
        }
    }
}
