namespace ClosedXML.Excel
{
    /// <summary>
    /// Indicates the type of rule being used to describe an area or aspect of the PivotTable.
    /// https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_PivotAreaType_topic_ID0E42RFB.html#topic_ID0E42RFB
    /// </summary>
    internal enum XLPivotStyleFormatRuleType
    {
        ///<summary>Refers to the whole PivotTable.</summary>
        None,
        ///<summary>Refers to a header or item.</summary>
        Normal,
        ///<summary>Refers to something in the data area.</summary>
        Data,
        ///<summary>Refers to the whole PivotTable.</summary>
        All,
        ///<summary>Refers to the blank cells at the top-left of the PivotTable (top-right for RTL sheets).</summary>
        Origin,
        ///<summary>Refers to a field button.</summary>
        Button,
        ///<summary>Refers to the blank cells at the top-right of the PivotTable (top-left for RTL sheets).</summary>
        TopRight,
        ///<summary>Refers to the blank cells at the top of the PivotTable, on its trailing edge (top-right for LTR sheets, top-left for RTL sheets).</summary>
        TopEnd
    }
}
