#nullable disable
namespace ClosedXML.Excel;

internal class XLPivotTableAxisFieldStyleFormats : IXLPivotFieldStyleFormats
{
    private readonly XLPivotTable _pivotTable;
    private readonly XLPivotTableAxisField _axisField;

    public XLPivotTableAxisFieldStyleFormats(XLPivotTable pivotTable, XLPivotTableAxisField axisField)
    {
        _pivotTable = pivotTable;
        _axisField = axisField;
    }

    #region IXLPivotFieldStyleFormats

    public IXLPivotValueStyleFormat DataValuesFormat { get; }

    public IXLPivotStyleFormat Header { get; }

    public IXLPivotStyleFormat Label
    {
        get
        {
            /* <x:pivotArea type="normal"
			 *              dataOnly="0"
			 *              labelOnly="1">
			 *     <x:references count="1">
			 *         <x:reference field="4"/>
			 *	   </x:references>
			 * </x:pivotArea>
             */
            return new XLPivotStyleFormat(
                _pivotTable,
                area => AreaBelongsToField(area, XLPivotStyleFormatElement.Label),
                () => CreateFieldArea(XLPivotStyleFormatElement.Label));
        }
    }

    public IXLPivotStyleFormat Subtotal { get; }

    #endregion

    private bool AreaBelongsToField(XLPivotArea area, XLPivotStyleFormatElement element)
    {
        if (area.References.Count != 1)
            return false;

        var field = area.References[0].Field;
        if (field is null || (int)field != _axisField.Offset)
            return false;

        return
            area.Field is null &&
            area.Type == XLPivotAreaType.Normal &&
            area.DataOnly == (element == XLPivotStyleFormatElement.Data) &&
            area.LabelOnly == (element == XLPivotStyleFormatElement.Label) &&
            !area.GrandRow &&
            !area.GrandCol &&
            area.CacheIndex == false &&
            area.Offset is null &&
            !area.CollapsedLevelsAreSubtotals &&
            area.Axis is null &&
            area.FieldPosition is null;
    }

    private XLPivotArea CreateFieldArea(XLPivotStyleFormatElement element)
    {
        var area = new XLPivotArea
        {
            DataOnly = (element == XLPivotStyleFormatElement.Data),
            LabelOnly = (element == XLPivotStyleFormatElement.Label),
        };
        area.AddReference(new XLPivotReference
        {
            Field = (uint)_axisField.Offset,
        });
        return area;
    }
}
