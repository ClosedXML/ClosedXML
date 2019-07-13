namespace ClosedXML.Excel
{
    // Rule describing a PivotTable selection.
    // Represented by <pivotArea> in OpenXML
    // https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_pivotArea_topic_ID0EQ4A5.html
    internal class XLPivotStyleFormatRule
    {
        internal const XLPivotStyleFormatRuleType DefaultRuleType = XLPivotStyleFormatRuleType.Data;
        internal bool CollapsedLevelsAreSubtotals { get; set; } = false;
        internal bool IsInOutlineMode { get; set; } = true;
        internal XLPivotStyleFormatRuleType RuleType { get; set; } = DefaultRuleType;
    }
}
