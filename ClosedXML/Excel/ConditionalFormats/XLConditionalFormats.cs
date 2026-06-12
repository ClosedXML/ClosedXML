using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel.CalcEngine.Visitors;
using ClosedXML.Parser;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A container for conditional formatting of a <see cref="XLWorksheet"/>. It contains
    /// a collection of <see cref="XLConditionalFormat"/>. Doesn't contain pivot table formats,
    /// they are in pivot table <see cref="XLPivotTable.ConditionalFormats"/>,
    /// </summary>
    internal class XLConditionalFormats : IXLConditionalFormats, IEnumerable<XLConditionalFormat>, ISheetListener
    {
        private readonly XLWorksheet _worksheet;
        private readonly List<XLConditionalFormat> _conditionalFormats = new();

        private static readonly List<XLConditionalFormatType> CFTypesExcludedFromConsolidation = new()
        {
            XLConditionalFormatType.DataBar,
            XLConditionalFormatType.ColorScale,
            XLConditionalFormatType.IconSet,
            XLConditionalFormatType.Top10,
            XLConditionalFormatType.AboveAverage,
            XLConditionalFormatType.IsDuplicate,
            XLConditionalFormatType.IsUnique
        };

        public XLConditionalFormats(XLWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public void Add(IXLConditionalFormat conditionalFormat)
        {
            var addedCf = (XLConditionalFormat)conditionalFormat;
            _conditionalFormats.Add(addedCf);
            if (addedCf.FormatValue is { } dxf)
                _worksheet.Workbook.Styles.RegisteredDxFormat(dxf);
        }

        public IEnumerator<XLConditionalFormat> GetEnumerator()
        {
            return _conditionalFormats.GetEnumerator();
        }

        IEnumerator<IXLConditionalFormat> IEnumerable<IXLConditionalFormat>.GetEnumerator()
        {
            return GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Remove(Predicate<IXLConditionalFormat> predicate)
        {
            _conditionalFormats.RemoveAll(predicate);
        }

        public void RemoveAll()
        {
            _conditionalFormats.Clear();
        }

        /// <summary>
        /// Reorders the according to original priority. Done during load process
        /// </summary>
        public void ReorderAccordingToOriginalPriority()
        {
            var reorderedFormats = _conditionalFormats.OrderBy(cf => cf.Priority).ToList();
            _conditionalFormats.Clear();
            _conditionalFormats.AddRange(reorderedFormats);
        }

        /// <summary>
        /// Clear conditional formats in the <paramref name="area"/>. Split if necessary, remove if
        /// conditional format has no area left.
        /// </summary>
        internal void Clear(Area area)
        {
            for (var i = _conditionalFormats.Count - 1; i >= 0; --i)
            {
                var conditionalFormat = _conditionalFormats[i];
                if (!conditionalFormat.Areas.IntersectsWith(area))
                    continue;

                var remainingAreas = conditionalFormat.Areas.Excluding(area);
                if (remainingAreas.Count > 0)
                {
                    conditionalFormat.Areas = remainingAreas;
                }
                else
                {
                    _conditionalFormats.RemoveAt(i);
                }
            }
        }

        internal void CopyFrom(XLWorksheet sourceSheet, Area sourceArea, Point targetPoint, bool mergeUncoveredInSameSheet = false)
        {
            // If source and target sheets are same, do not go over the end
            var sourceCfCount = sourceSheet.ConditionalFormats._conditionalFormats.Count;
            for (var i = 0; i < sourceCfCount; ++i)
            {
                var sourceCf = sourceSheet.ConditionalFormats._conditionalFormats[i];
                if (!sourceCf.Areas.TryCopyAreaTo(targetPoint, sourceArea, out var targetAreas))
                    continue;

                // Legacy behavior where a copied single point was merged into CF when not covered.
                // But only for cell copy API, nor range copy API (even if range is only 1x1).
                if (mergeUncoveredInSameSheet && _worksheet == sourceSheet)
                {
                    foreach (var targetArea in targetAreas)
                    {
                        var isCovered = sourceCf.Areas.Any(sourceCfArea => sourceCfArea.Covers(targetArea));
                        if (!isCovered)
                        {
                            sourceCf.Areas = sourceCf.Areas.With(targetArea);
                        }
                    }
                }
                else
                {
                    var targetCfCopy = new XLConditionalFormat(_worksheet, sourceCf, targetAreas);
                    Add(targetCfCopy);
                }
            }
        }

        /// <summary>
        /// The method consolidate the same conditional formats, which are located in adjacent ranges.
        /// </summary>
        internal void Consolidate()
        {
            var formats = _conditionalFormats
                .Where(cf => cf.Ranges.Any())
                .ToList();
            _conditionalFormats.Clear();

            while (formats.Count > 0)
            {
                var format = formats[0];
                if (!CFTypesExcludedFromConsolidation.Contains(format.ConditionalFormatType))
                {
                    var (rulesToConsolidate, areasWithSameFormat) = GetConsolidatableRules(formats);
                    var consolidatedAreas = areasWithSameFormat.GetConsolidated();
                    var consolidatedCf = new XLConditionalFormat(_worksheet, format, consolidatedAreas);

                    // Remove consolidated formats
                    rulesToConsolidate.Reverse();
                    foreach (var consolidatedRuleIndex in rulesToConsolidate)
                        formats.RemoveAt(consolidatedRuleIndex);

                    format = consolidatedCf;
                }

                _conditionalFormats.Add(format);
                formats.RemoveAt(0);
            }
        }

        private static (List<int> RulesToConsolidate, XLAreaList AreaList) GetConsolidatableRules(List<XLConditionalFormat> conditionalFormats)
        {
            var rule = conditionalFormats[0];
            var sameFormatAreas = rule.Areas.ToList();
            var differentFormatAreas = new List<Area>();

            // The ids to the list must be in the ascending order
            var rulesToConsolidate = new List<int>();
            for (int i = 1; i < conditionalFormats.Count; ++i)
            {
                var candidateRule = conditionalFormats[i];

                var intersectsDifferentFormatAreas = differentFormatAreas.Any(differentFormatArea => candidateRule.Areas.Any(v => v.Intersects(differentFormatArea)));
                if (intersectsDifferentFormatAreas)
                {
                    // We reached a rule intersecting any of captured ranges. Stop for not breaking the priorities.
                    break;
                }

                var isSameFormat = XLConditionalFormat.NoRangeComparer.Equals(candidateRule, rule);
                if (isSameFormat)
                {
                    // We reached a rule that has same format as the consolidated rule and doesn't intersect different
                    // format areas. We can consolidate the candidate rule with the rule without potentially breaking
                    // any rule with a priority between rule and candidate rule.
                    sameFormatAreas.AddRange(candidateRule.Areas);
                    rulesToConsolidate.Add(i);
                    continue;
                }

                var intersectsSameFormatAreas = sameFormatAreas.Any(sameFormatArea => candidateRule.Areas.Any(v => v.Intersects(sameFormatArea)));
                if (intersectsSameFormatAreas)
                {
                    // We reached a rule that has different format and intersects area to be consolidated. That means
                    // it's not possible to consolidate any subsequent rule, because it could break this one, and
                    // consolidation must stop here.
                    break;
                }

                // The most common case: The candidate rule has a different format and doesn't intersect the sameFormatAreas
                // The format thus must be added to the differentFormatAreas, because it can interrupt subsequent rules.
                differentFormatAreas.AddRange(candidateRule.Areas);
            }

            return (rulesToConsolidate, new XLAreaList(sameFormatAreas));
        }

        #region ISheetListener

        void ISheetListener.OnInsertAreaAndShiftDown(XLWorksheet sheet, Area insertedArea)
        {
            var inserted = new XLBookArea(sheet.Name, insertedArea);
            var refMod = new ReferenceShiftOnInsertRefModVisitor(inserted, true);
            AdjustFormulas(refMod);

            AdjustConditionalFormatAreas(sheet, inserted.Area, static (sqref, insertedArea) => sqref.InsertAndShiftDown(insertedArea));
        }

        void ISheetListener.OnInsertAreaAndShiftRight(XLWorksheet sheet, Area insertedArea)
        {
            var inserted = new XLBookArea(sheet.Name, insertedArea);
            var refMod = new ReferenceShiftOnInsertRefModVisitor(inserted, false);
            AdjustFormulas(refMod);

            AdjustConditionalFormatAreas(sheet, inserted.Area, static (sqref, insertedArea) => sqref.InsertAndShiftRight(insertedArea));
        }

        void ISheetListener.OnDeleteAreaAndShiftLeft(XLWorksheet sheet, Area deletedArea)
        {
            var deleted = new XLBookArea(sheet.Name, deletedArea);
            var refMod = new ReferenceShiftOnDeleteRefModVisitor(deleted, XLShiftDeletedCells.ShiftCellsLeft);
            AdjustFormulas(refMod);

            AdjustConditionalFormatAreas(sheet, deleted.Area, static (sqref, deletedArea) => sqref.DeleteAndShiftLeft(deletedArea));
        }

        void ISheetListener.OnDeleteAreaAndShiftUp(XLWorksheet sheet, Area deletedArea)
        {
            var deleted = new XLBookArea(sheet.Name, deletedArea);
            var refMod = new ReferenceShiftOnDeleteRefModVisitor(deleted, XLShiftDeletedCells.ShiftCellsUp);
            AdjustFormulas(refMod);

            AdjustConditionalFormatAreas(sheet, deleted.Area, static (sqref, deletedArea) => sqref.DeleteAndShiftUp(deletedArea));
        }

        private void AdjustFormulas(CopyVisitor refMod)
        {
            foreach (var conditionalFormat in _conditionalFormats)
            {
                var anchor = conditionalFormat.Areas[0].FirstPoint;
                var formulaIndexes = conditionalFormat.Values.Where(x => x.Value.IsFormula).Select(x => x.Key).ToArray();
                foreach (var index in formulaIndexes)
                {
                    var originalFormula = conditionalFormat.Values[index];
                    var shiftedFormula = FormulaConverter.ModifyA1(originalFormula.Value, _worksheet.Name, anchor.Row, anchor.Column, refMod);
                    conditionalFormat.Values[index] = new XLFormula(shiftedFormula) { IsFormula = true };
                }
            }
        }

        private void AdjustConditionalFormatAreas(XLWorksheet sheet, Area affectedRange, Func<XLAreaList, Area, XLAreaList> adjustAreas)
        {
            if (sheet != _worksheet)
                return;

            for (var i = _conditionalFormats.Count - 1; i >= 0; --i)
            {
                var conditionalFormat = _conditionalFormats[i];
                var modifiedAreaList = adjustAreas(conditionalFormat.Areas, affectedRange);
                if (modifiedAreaList.Count == 0)
                {
                    _conditionalFormats.RemoveAt(i);
                }
                else
                {
                    conditionalFormat.Areas = modifiedAreaList;
                }
            }
        }

        #endregion
    }
}
