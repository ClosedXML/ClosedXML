using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A container for conditional formatting of a <see cref="XLWorksheet"/>. It contains
    /// a collection of <see cref="XLConditionalFormat"/>. Doesn't contain pivot table formats,
    /// they are in pivot table <see cref="XLPivotTable.ConditionalFormats"/>,
    /// </summary>
    internal class XLConditionalFormats : IXLConditionalFormats, IEnumerable<XLConditionalFormat>
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
                    var originalAnchor = format.Areas[0].FirstPoint;

                    var (rulesToConsolidate, areasWithSameFormat) = GetConsolidateableRules(formats);
                    var consolidatedAreas = areasWithSameFormat.GetConsolidated();
                    var consolidatedCf = new XLConditionalFormat(_worksheet, format, consolidatedAreas);

                    // Remove consolidated formats
                    rulesToConsolidate.Reverse();
                    foreach (var consolidatedRuleIndex in rulesToConsolidate)
                        formats.RemoveAt(consolidatedRuleIndex);

                    var consolidatedAnchor = consolidatedAreas[0].FirstPoint;
                    consolidatedCf.AdjustFormulas(_worksheet.Cell(originalAnchor), _worksheet.Cell(consolidatedAnchor));
                    format = consolidatedCf;
                }

                _conditionalFormats.Add(format);
                formats.RemoveAt(0);
            }
        }

        private static (List<int> RulesToConsolidate, XLAreaList AreaList) GetConsolidateableRules(List<XLConditionalFormat> conditionalFormats)
        {
            var rule = conditionalFormats[0];
            var sameFormatAreas = rule.Areas.ToList();
            var differentFormatAreas = new List<XLSheetRange>();

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
                    // We reached a rule that has same format as the condolidated rule and doesn't interect different
                    // format areas. We can consolidate the candidate rule with the rule without potentially breaking
                    // any rule with a priority between rule and candidate rule.
                    sameFormatAreas.AddRange(candidateRule.Areas);
                    rulesToConsolidate.Add(i);
                    continue;
                }

                var intersectsSameFormatAreas = sameFormatAreas.Any(sameFormatArea => candidateRule.Areas.Any(v => v.Intersects(sameFormatArea)));
                if (intersectsSameFormatAreas)
                {
                    // We reached a rule that has differnet format and intersects area to be consolidated. That means
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
    }
}
