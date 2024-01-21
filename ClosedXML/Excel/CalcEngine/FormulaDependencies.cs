using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A list of objects a cell formula depends on. If one of them changes,
    /// the formula value might no longer be accurate and needs to be recalculated.
    /// </summary>
    internal class FormulaDependencies
    {
        private readonly HashSet<XLBookArea> _areas = new();
        private readonly HashSet<XLName> _names = new();

        /// <summary>
        /// List of areas the formula depends on. It is likely a superset of accurate
        /// result for unusual formulas, but if a value in an areas changes, the dependent
        /// formula should be marked as dirty.
        /// </summary>
        public IReadOnlyCollection<XLBookArea> Areas => _areas;

        /// <summary>
        /// A collection of names in the formula. If a name changes (added, deleted),
        /// the formula dependencies should be refreshed, because new name might refer to
        /// different references (e.g. a name previously referred to <c>A5</c> and is redefined
        /// to <c>B7</c> or just value <c>7</c> =&gt; formula no longer depends on <c>A5</c>).
        /// </summary>
        public IReadOnlyCollection<XLName> Names => _names;

        internal void AddAreas(List<XLBookArea> sheetAreas)
        {
            _areas.UnionWith(sheetAreas);
        }

        internal void AddName(XLName name)
        {
            _names.Add(name);
        }

        internal void RenameSheet(string oldSheetName, string newSheetName)
        {
            // The renaming is done for every formula, so only allocate when needed.
            List<(XLBookArea Original, XLBookArea Replacement)>? areasToRename = null;
            foreach (var areaInFormula in _areas)
            {
                if (XLHelper.SheetComparer.Equals(areaInFormula.Name, oldSheetName))
                {
                    var renamedArea = new XLBookArea(newSheetName, areaInFormula.Area);
                    areasToRename ??= new List<(XLBookArea Original, XLBookArea Replacement)>();
                    areasToRename.Add((areaInFormula, renamedArea));
                }
            }

            if (areasToRename is not null)
            {
                foreach (var (original, replacement) in areasToRename)
                {
                    _areas.Remove(original);
                    _areas.Add(replacement);
                }
            }

            List<(XLName Original, XLName Replacement)>? namesToRename = null;
            foreach (var nameInFormula in _names)
            {
                if (nameInFormula.SheetName is not null &&
                    XLHelper.SheetComparer.Equals(nameInFormula.SheetName, oldSheetName))
                {
                    var renamedName = new XLName(newSheetName, nameInFormula.Name);
                    namesToRename ??= new List<(XLName Original, XLName Replacement)>();
                    namesToRename.Add((nameInFormula, renamedName));
                }
            }

            if (namesToRename is not null)
            {
                foreach (var (original, replacement) in namesToRename)
                {
                    _names.Remove(original);
                    _names.Add(replacement);
                }
            }
        }
    }
}
